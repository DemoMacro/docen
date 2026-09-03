import {
  argbHex,
  decodeObject,
  type BrushInfo,
  type ImageInfo,
  type PathInfo,
  type PenInfo,
} from "./decode";
import type { Draft, PathCmds, PathDraft, PicDraft } from "./draft";
import { pushPath, pushRect } from "./draft";
import { boxOf, sameBox, sameSourceBlend } from "./finalize";
import { gdiTextDrafts } from "./gdi";
import {
  GDIPLUS_VERSION,
  PLUS_DRAW_IMAGE_POINTS,
  PLUS_DRAW_LINES,
  PLUS_DRAW_PATH,
  PLUS_END_OF_FILE,
  PLUS_FILL_PATH,
  PLUS_FILL_RECTS,
  PLUS_MULTIPLY_WORLD_TRANSFORM,
  PLUS_OBJECT,
  PLUS_RESET_WORLD_TRANSFORM,
  PLUS_RESTORE,
  PLUS_SAVE,
  PLUS_SET_WORLD_TRANSFORM,
} from "./records";
import { combine, IDENTITY, readXform, xformPoint, type Xform } from "./xform";

/** Nested-metafile recursion guard — a self-referencing or malformed blob
 *  must not spin through unbounded container levels. */
const MAX_NESTING = 3;

/** All drawable drafts of a raw carrier EMF: its concatenated EMF+ comment
 *  stream replayed with `basis` pre-applied on top of every world transform,
 *  plus the GDI-side text chain. `basis` carries nested DrawImagePoints
 *  parallelograms: the parent maps the image rect onto it and the child plays
 *  under that mapping. */
export function carrierDrafts(emf: Uint8Array, basis: Xform | undefined, depth: number): Draft[] {
  // Concatenate every EMR_GDICOMMENT "EMF+" payload into one record stream.
  const ev = new DataView(emf.buffer, emf.byteOffset, emf.byteLength);
  let plusLen = 0;
  for (let eo = 0; eo + 8 <= emf.length;) {
    const type = ev.getUint32(eo, true);
    const size = ev.getUint32(eo + 4, true);
    if (size < 8 || eo + size > emf.length) break;
    if (type === 70 && ev.getUint32(eo + 12, true) === 0x2b464d45) plusLen += size - 16;
    eo += size;
  }
  if (!plusLen) return [];
  const plus = new Uint8Array(plusLen);
  let pw = 0;
  for (let eo = 0; eo + 8 <= emf.length;) {
    const type = ev.getUint32(eo, true);
    const size = ev.getUint32(eo + 4, true);
    if (size < 8 || eo + size > emf.length) break;
    if (type === 70 && ev.getUint32(eo + 12, true) === 0x2b464d45) {
      plus.set(emf.subarray(eo + 16, eo + size), pw);
      pw += size - 16;
    }
    eo += size;
  }
  // Every draft coordinate passes through this effective transform first. A
  // nested level applies its own world FIRST (record space → nested device)
  // and THEN the parent's image-to-parallelogram basis (device → placed box):
  // combine is "left first", so the order here must be (xf, basis).
  const effOf = (xf: Xform): Xform => (basis ? combine(xf, basis) : xf);

  const pv = new DataView(plus.buffer);
  const objects = new Map<number, unknown>();
  let xf: Xform = { ...IDENTITY };
  const saved: Xform[] = [];
  const drafts: Draft[] = [];
  let lastBrush: BrushInfo | undefined;
  let lastPen: PenInfo | undefined;

  // Object payloads are [chunkDataSize u32][versionStamp-or-bytes]: complete
  // definitions start with the GDI+ version stamp, continuation chunks carry
  // raw object bytes with no stamp. Giant images arrive as same-slot record
  // runs whose chunk flag keeps saying "continued" even on the last chunk,
  // and several different-slot definitions may sit open at once before one
  // consuming draw. Open runs therefore live in a per-slot table and install
  // when a consuming record arrives.
  type OpenRun = { type: number; parts: Uint8Array[] };
  const openRuns = new Map<number, OpenRun>();
  const installRun = (slot: number, run: OpenRun): void => {
    let payload = run.parts[0];
    if (run.parts.length > 1) {
      const total = run.parts.reduce((n, p) => n + p.length, 0);
      const all = new Uint8Array(total);
      let at = 0;
      for (const part of run.parts) {
        all.set(part, at);
        at += part.length;
      }
      payload = all;
    }
    installObject(objects, slot, decodeObject(payload, run.type));
  };
  const installOpenRuns = (): void => {
    if (!openRuns.size) return;
    for (const [slot, run] of openRuns) installRun(slot, run);
    openRuns.clear();
  };
  const installObject = (
    table: Map<number, unknown>,
    slot: number,
    obj: BrushInfo | PenInfo | PathInfo | ImageInfo | undefined,
  ): void => {
    if (!obj) return;
    if ("solid" in obj) lastBrush = obj;
    else if ("width" in obj) lastPen = obj;
    table.set(slot, obj);
  };

  for (let po = 0; po + 8 <= plus.length;) {
    const rt = pv.getUint16(po, true);
    const flags = pv.getUint16(po + 2, true);
    const rs = pv.getUint32(po + 4, true);
    if (rs < 8 || po + rs > plus.length) {
      installOpenRuns();
      break;
    }
    if (rt === PLUS_END_OF_FILE) {
      installOpenRuns();
      break;
    }
    const d = po + 8;
    const objectId = flags & 0xff;
    switch (rt) {
      case PLUS_OBJECT: {
        // The version stamp check reads d+4..8 — a record that small cannot
        // carry even the object header, so skip it.
        if (rs < 16) break;
        const complete =
          (pv.getUint32(d + 4, true) & 0xffff0000) === (GDIPLUS_VERSION & 0xffff0000);
        const run = openRuns.get(objectId);
        if (run && !complete) {
          // Continuation chunk — same two-word header as the opening record
          // ([chunk size][object total]) followed by raw object bytes.
          run.parts.push(plus.subarray(d + 8, po + rs));
          break;
        }
        if (run) installRun(objectId, run); // fresh same-slot definition
        openRuns.delete(objectId);
        openRuns.set(objectId, { type: (flags >> 8) & 0x7f, parts: [plus.subarray(d, po + rs)] });
        break;
      }
      case PLUS_SAVE:
        saved.push(xf);
        xf = { ...xf };
        break;
      case PLUS_RESTORE:
        xf = saved.pop() ?? xf;
        break;
      case PLUS_SET_WORLD_TRANSFORM:
        xf = readXform(pv, d);
        break;
      case PLUS_RESET_WORLD_TRANSFORM:
        xf = { ...IDENTITY };
        break;
      case PLUS_MULTIPLY_WORLD_TRANSFORM:
        xf = combine(xf, readXform(pv, d));
        break;
      case PLUS_FILL_PATH: {
        installOpenRuns();
        const path = objects.get(objectId) as PathInfo | undefined;
        // ColorEmphasis flag (0x8000): the record carries its own solid color
        // inline ([u32][ARGB] behind the header) instead of the last-defined
        // brush — corpus census shows every inlined color is opaque ARGB.
        const fill = flags & 0x8000 ? argbHex(pv.getUint32(d + 4, true)) : lastBrush?.solid;
        if (path?.cmds && fill) pushPath(drafts, path.cmds, effOf(xf), { fill });
        break;
      }
      case PLUS_DRAW_PATH: {
        installOpenRuns();
        const path = objects.get(objectId) as PathInfo | undefined;
        if (path?.cmds && lastPen?.color) {
          // A World-unit pen's width rides its world transform like text does
          // (the same rotation-safe column-norm factor gdiTextDrafts applies).
          const eff = effOf(xf);
          const scale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
          pushPath(drafts, path.cmds, eff, {
            strokeColor: lastPen.color,
            strokeWidth: lastPen.width * scale,
            ...(lastPen.dash ? { dash: lastPen.dash } : {}),
          });
        }
        break;
      }
      case PLUS_FILL_RECTS: {
        installOpenRuns();
        // Same corpus conventions as FillPath above: the WMF-embedded record
        // prepends a payload-size u32 (the OBJECT chunkDataSize header), so an
        // inlined ColorEmphasis ARGB sits at d+4 and the rect list follows.
        // 0x4000 marks compressed EmfPlusRectS rects (int16, 8 bytes).
        const fill = flags & 0x8000 ? argbHex(pv.getUint32(d + 4, true)) : lastBrush?.solid;
        if (rs < 12) break;
        const n = pv.getUint32(d + 8, true);
        const step = flags & 0x4000 ? 8 : 16;
        // rs counts the record's 8-byte header — the rect list must end by
        // d + rs - 8 (= po + rs), not d + rs.
        if (!n || n > 10_000 || d + 12 + n * step > d + rs - 8) break;
        const eff = effOf(xf);
        for (let i = 0; i < n; i++) {
          const base = d + 12 + i * step;
          const rx = step === 8 ? pv.getInt16(base, true) : pv.getFloat32(base, true);
          const ry = step === 8 ? pv.getInt16(base + 2, true) : pv.getFloat32(base + 4, true);
          const rw = step === 8 ? pv.getInt16(base + 4, true) : pv.getFloat32(base + 8, true);
          const rh = step === 8 ? pv.getInt16(base + 6, true) : pv.getFloat32(base + 12, true);
          const [x, y] = xformPoint(eff, rx, ry);
          const [x2, y2] = xformPoint(eff, rx + rw, ry + rh);
          pushRect(drafts, x, y, x2 - x, y2 - y, fill);
        }
        break;
      }
      case PLUS_DRAW_LINES: {
        installOpenRuns();
        // Corpus conventions as above: the WMF-embedded record prepends its
        // payload-size u32, so the point count sits at d+4 and the polyline
        // points follow (0x8000 = compressed PointS int16 pairs, else float
        // PointF). The stroke comes from the object-table pen the record's
        // ObjectId selects — the list-page underlines ride these records.
        const pen = objects.get(objectId) as PenInfo | undefined;
        const color = pen?.color ?? lastPen?.color;
        if (!color) break;
        if (rs < 12) break;
        const n = pv.getUint32(d + 4, true);
        const compressed = (flags & 0x8000) !== 0;
        const step = compressed ? 4 : 8;
        // rs includes the 8-byte header (same accounting as FILL_RECTS).
        if (!n || n > 10_000 || d + 8 + n * step > d + rs - 8) break;
        const eff = effOf(xf);
        const cmds: PathCmds = [];
        for (let i = 0; i < n; i++) {
          const p = d + 8 + i * step;
          const px = compressed ? pv.getInt16(p, true) : pv.getFloat32(p, true);
          const py = compressed ? pv.getInt16(p + 2, true) : pv.getFloat32(p + 4, true);
          cmds.push([i === 0 ? "M" : "L", [px, py]]);
        }
        const scale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
        pushPath(drafts, cmds, eff, {
          strokeColor: color,
          strokeWidth: Math.max((pen?.width ?? 1) * scale, 1),
          ...(pen?.dash ? { dash: pen.dash } : {}),
        });
        break;
      }
      case PLUS_DRAW_IMAGE_POINTS: {
        installOpenRuns();
        const img = objects.get(objectId) as ImageInfo | undefined;
        if (!img) break;
        if (img.emfBytes) {
          drawNestedImage(img.emfBytes, pv, d, rs, flags, effOf(xf), depth, drafts);
        } else if (img.src) {
          drawImagePoints(pv, plus, d, rs, flags, effOf(xf), img, drafts);
        }
        break;
      }
      default:
        break;
    }
    po += rs;
  }
  // Dual-mode carriers hold two complete encodings of the same art: the EMF+
  // records an EMF+ player replays, and the GDI records kept as the fallback
  // for legacy players ([MS-EMFPLUS] §3.1.1; RFC 7903: either set alone
  // reconstitutes the drawing). A GDI picture would stack the coarser raster
  // over the vector art — photos re-blitted whole, halftone textures laid
  // across the composed design, white-backed masks where the EMF+ layer is
  // already transparent — so once the EMF+ layer pictures anything, GDI blts
  // drop entirely; they survive only when the EMF+ layer emits no pictures
  // (a bare EMF+ header with all imagery on the GDI side). GDI paths drop
  // only against a same-source EMF+ paint: same box, and the EMF+ fill is
  // the GDI flat COLORREF composited over white at one uniform alpha — the
  // systematic relation of a dual pair, since argbHex pre-blends partial
  // alpha (an EMF+ 25%-alpha gold reads FFF9E5 against a GDI FFE699). Art
  // that lives only on the GDI side (corpus wavy panels whose EMF+
  // counterparts draw fragments) keeps its distinct box or unrelated color
  // and survives. Text never drops: the EMF+ walk does not decode DrawString
  // runs yet, and a dual file's GDI text is the same glyphs at the same
  // spots.
  const plusPaints = drafts
    .filter((dr): dr is PathDraft => dr.kind === "path" && (!!dr.fill || !!dr.strokeColor))
    .map((dr) => ({ fill: dr.fill, box: boxOf(dr) }));
  const plusHasPic = drafts.some((dr) => dr.kind === "pic");
  const gdi = gdiTextDrafts(emf, effOf);
  drafts.push(
    ...gdi.filter((g) => {
      if (g.kind === "pic") return !plusHasPic;
      if (g.kind === "path") {
        const { fill } = g;
        const box = boxOf(g);
        return !(
          fill != null &&
          plusPaints.some((p) => sameBox(p.box, box) && sameSourceBlend(p.fill, fill))
        );
      }
      return true;
    }),
  );
  return drafts;
}

/** One metafile-typed EmfPlusImage drawn onto a destination parallelogram:
 *  like the bitmap arm, the record's source rectangle selects a sub-region
 *  of the nested metafile's device space — that sub-domain maps onto the
 *  corners and the nested replay runs under it as its basis. */
function drawNestedImage(
  nested: Uint8Array,
  view: DataView,
  d: number,
  rs: number,
  flags: number,
  outerEff: Xform,
  depth: number,
  drafts: Draft[],
): void {
  if (depth >= MAX_NESTING) return;
  // rs includes the 8-byte record header (same accounting as FILL_RECTS).
  const end = Math.min(view.byteLength, d + rs - 8);
  const pts = d + 32;
  if (pts + 12 > end || view.getUint32(d + 28, true) !== 3) return;
  const srcX = view.getFloat32(d + 12, true);
  const srcY = view.getFloat32(d + 16, true);
  const srcW = view.getFloat32(d + 20, true);
  const srcH = view.getFloat32(d + 24, true);
  if (!Number.isFinite(srcW) || !Number.isFinite(srcH) || srcW <= 0 || srcH <= 0) return;
  const compressed = (flags & 0x4000) !== 0;
  const read = (i: number): [number, number] => {
    const p = pts + i * (compressed ? 4 : 8);
    return compressed
      ? [view.getInt16(p, true), view.getInt16(p + 2, true)]
      : [view.getFloat32(p, true), view.getFloat32(p + 4, true)];
  };
  const corners = [read(0), read(1), read(2)].map(([x, y]) => xformPoint(outerEff, x, y));
  const [p0, p1, p2] = corners;
  // Affine taking the source sub-domain onto the drawn parallelogram.
  const basis: Xform = {
    m11: (p1[0] - p0[0]) / srcW,
    m21: (p1[1] - p0[1]) / srcW,
    m12: (p2[0] - p0[0]) / srcH,
    m22: (p2[1] - p0[1]) / srcH,
    dx: p0[0] - srcX * ((p1[0] - p0[0]) / srcW),
    dy: p0[1] - srcY * ((p2[1] - p0[1]) / srcH),
  };
  drafts.push(...carrierDrafts(nested, basis, depth + 1));
}

function drawImagePoints(
  view: DataView,
  buf: Uint8Array,
  d: number,
  rs: number,
  flags: number,
  xf: Xform,
  img: ImageInfo,
  drafts: Draft[],
): void {
  // Payload layout (corpus-verified against the [MS-EMFPLUS] field table):
  // [attrsId u32][srcUnit u32][reserved u32][SrcRect RectF][count u32][points].
  // The source rectangle is ALWAYS present and selects a sub-region of the
  // bitmap — Office sprite-sheets entire pages of art into one image object,
  // so ignoring it smears the sheet over every destination box.
  const end = Math.min(buf.length, d + rs - 8);
  const pts = d + 32;
  if (pts + 12 > end || view.getUint32(d + 28, true) !== 3) return;
  const srcX = view.getFloat32(d + 12, true);
  const srcY = view.getFloat32(d + 16, true);
  const srcW = view.getFloat32(d + 20, true);
  const srcH = view.getFloat32(d + 24, true);
  const compressed = (flags & 0x4000) !== 0;
  const read = (i: number): [number, number] => {
    const p = pts + i * (compressed ? 4 : 8);
    return compressed
      ? [view.getInt16(p, true), view.getInt16(p + 2, true)]
      : [view.getFloat32(p, true), view.getFloat32(p + 4, true)];
  };
  const corners = [read(0), read(1), read(2)].map(([x, y]) => xformPoint(xf, x, y));
  const xs = corners.map((c) => c[0]);
  const ys = corners.map((c) => c[1]);
  const minX = Math.min(...xs),
    minY = Math.min(...ys);
  const w = Math.max(...xs) - minX,
    h = Math.max(...ys) - minY;
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return;
  if (!Number.isFinite(srcW) || !Number.isFinite(srcH) || srcW <= 0 || srcH <= 0) return;
  let crop: PicDraft["crop"] | undefined =
    img.w && img.h
      ? {
          l: srcX / img.w,
          t: srcY / img.h,
          r: Math.max(0, 1 - (srcX + srcW) / img.w),
          b: Math.max(0, 1 - (srcY + srcH) / img.h),
        }
      : undefined;
  // A rect that still covers the whole bitmap must not engage the renderer's
  // offscreen-copy path — exporters emit degenerate full-cover rects freely.
  if (crop && !(crop.l > 0.001 || crop.t > 0.001 || crop.r > 0.001 || crop.b > 0.001)) {
    crop = undefined;
  }
  drafts.push({ kind: "pic", x: minX, y: minY, w, h, src: img.src!, ...(crop ? { crop } : {}) });
}
