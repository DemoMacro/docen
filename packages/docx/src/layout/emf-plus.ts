// Office renders clipboard/import vector media as dual-mode metafiles: the
// placeable-WMF body carries only a coarse GDI approximation, while
// META_ESCAPE chunks (magic "WMFC") embed the real drawing as a complete EMF
// whose EMR_GDICOMMENT payloads hold the GDI+ (EMF+) record stream. This
// module reassembles that stream and replays it into the same structured
// drawing members native OOXML drawings are projected into, so metafile art
// renders through the identical layout/paint model instead of a raster detour.

import type { LayoutDrawingMember } from "@docen/layout";

const PLACEABLE_MAGIC = 0x9ac6cdd7;
const META_ESCAPE = 0x0626;
const WMFC_MAGIC = 0x43464d57; // "WMFC"
const WMFC_CHUNK_HEADER = 34;
/** Every GDI+ object payload in these files repeats this version stamp. */
const GDIPLUS_VERSION = 0xdbc01002;

/** Reassemble the nested EMF carried by the WMFC escape chunks, or undefined
 *  when the bytes are not such a WMF. */
export function embeddedEmfStream(bytes: Uint8Array): Uint8Array | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 44 || view.getUint32(0, true) !== PLACEABLE_MAGIC) return undefined;
  const chunks: Uint8Array[] = [];
  let off = 40; // placeable + standard WMF header
  while (off + 6 <= bytes.length) {
    const sizeWords = view.getUint32(off, true);
    const fn = view.getUint16(off + 4, true);
    if (sizeWords < 3) break;
    const end = off + sizeWords * 2;
    if (end > bytes.length) break;
    if (fn === META_ESCAPE && view.getUint32(off + 10, true) === WMFC_MAGIC) {
      const cb = Math.min(view.getUint16(off + 8, true), end - off - 10);
      chunks.push(bytes.subarray(off + 10 + WMFC_CHUNK_HEADER, off + 10 + cb));
    }
    off = end;
  }
  if (!chunks.length) return undefined;
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const emf = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) {
    emf.set(c, o);
    o += c.length;
  }
  return total >= 108 ? emf : undefined;
}

// ── record codes replayed ──

const PLUS_END_OF_FILE = 0x4002;
const PLUS_OBJECT = 0x4008;
const PLUS_FILL_RECTS = 0x400a;
const PLUS_FILL_PATH = 0x4014;
const PLUS_DRAW_PATH = 0x4015;
const PLUS_DRAW_IMAGE_POINTS = 0x401b;
const PLUS_SAVE = 0x4025;
const PLUS_RESTORE = 0x4026;
const PLUS_SET_WORLD_TRANSFORM = 0x402a;
const PLUS_RESET_WORLD_TRANSFORM = 0x402b;
const PLUS_MULTIPLY_WORLD_TRANSFORM = 0x402c;

const OBJ_BRUSH = 1;
const OBJ_PEN = 2;
const OBJ_PATH = 3;
const OBJ_IMAGE = 5;

interface Xform {
  m11: number;
  m12: number;
  m21: number;
  m22: number;
  dx: number;
  dy: number;
}

const IDENTITY: Xform = { m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 };

function xformPoint(t: Xform, x: number, y: number): [number, number] {
  return [x * t.m11 + y * t.m21 + t.dx, y * t.m22 + x * t.m12 + t.dy];
}

function combine(a: Xform, b: Xform): Xform {
  return {
    m11: a.m11 * b.m11 + a.m21 * b.m12,
    m12: a.m12 * b.m11 + a.m22 * b.m12,
    m21: a.m11 * b.m21 + a.m21 * b.m22,
    m22: a.m12 * b.m21 + a.m22 * b.m22,
    dx: a.dx * b.m11 + a.dy * b.m12 + b.dx,
    dy: a.dx * b.m21 + a.dy * b.m22 + b.dy,
  };
}

/** XFORM payload: some exporters prefix a byte-length word (0x18); accept
 *  both shapes so a bare matrix still parses. */
function readXform(view: DataView, at: number): Xform {
  const base = view.getUint32(at, true) === 24 ? at + 4 : at;
  return {
    m11: view.getFloat32(base, true),
    m12: view.getFloat32(base + 4, true),
    m21: view.getFloat32(base + 8, true),
    m22: view.getFloat32(base + 12, true),
    dx: view.getFloat32(base + 16, true),
    dy: view.getFloat32(base + 20, true),
  };
}

/** ARGB word (0xAARRGGBB) → opaque hex RRGGBB; undefined when transparent. */
function argbHex(word: number): string | undefined {
  return word >>> 24 === 0 ? undefined : ((word & 0xffffff) | 0x1000000).toString(16).slice(1);
}

// A sub-path as raw command tuples; "M" starts, "L"/"C" continue, "Z" closes.
type PathCmds = Array<["M" | "L" | "C" | "Z", number[]]>;

interface PathDraft {
  kind: "path";
  cmds: PathCmds;
  fill?: string;
  strokeWidth?: number;
  strokeColor?: string;
}

interface PicDraft {
  kind: "pic";
  x: number;
  y: number;
  w: number;
  h: number;
  src: string;
  /** Source-rectangle crop, fractions of each image edge (DrawImagePoints). */
  crop?: { l: number; t: number; r: number; b: number };
}

/** A GDI-side ExtTextOutW run, kept in the same world coordinate space as
 *  the EMF+ drafts so finalize maps everything through one normalization.
 *  Dual-mode files carry real text as GDI records — their EmfPlusDrawString
 *  counterparts are zero-data stubs. */
interface TextDraft {
  kind: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  family: string;
  sizeWorld: number;
  color?: string;
  bold?: boolean;
}

type Draft = PathDraft | PicDraft | TextDraft;

// ── object decoding ──
// An object's payload starts at its TotalObjectSize word ([0..3]), then the
// shared version stamp ([4..7]); per-type fields begin at [+8]. Large images
// span multiple consecutive EmfPlusObject records — the caller assembles the
// full byte set first.

interface BrushInfo {
  solid?: string;
}
interface PenInfo {
  width: number;
  color?: string;
}
interface PathInfo {
  cmds: PathCmds;
}

function decodeObject(
  payload: Uint8Array,
  type: number,
): BrushInfo | PenInfo | PathInfo | ImageInfo | undefined {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const end = payload.length;
  // Payloads start [TotalObjectSize][Version] with type fields at +8; a
  // re-assembled chunk run keeps the exporter's [chunkDataSize][objectTotal]
  // prefix in front of the stamp — locate the stamp to find the field origin.
  let base = 8;
  if (
    view.getUint32(4, true) !== GDIPLUS_VERSION &&
    end >= 16 &&
    view.getUint32(8, true) === GDIPLUS_VERSION
  )
    base = 12;
  switch (type) {
    case OBJ_BRUSH: {
      const brushType = view.getUint32(base, true);
      if (brushType === 0) return { solid: argbHex(view.getUint32(base + 4, true)) };
      // PathGradient brushes take their declared center color as the flat
      // approximation — the member protocol paints solids only, so a bounded
      // scan for the first bright opaque word stands in until gradient fills
      // become a member-level concept.
      for (let p = base + 4; p < Math.min(end - 3, base + 144); p += 4) {
        const c = view.getUint32(p, true);
        if (c >>> 24 !== 0xff) continue;
        const rgb = c & 0xffffff;
        const lum = ((rgb >> 16) & 0xff) * 299 + ((rgb >> 8) & 0xff) * 587 + (rgb & 0xff) * 114;
        if (lum > 120_000) return { solid: argbHex(c) };
      }
      return undefined;
    }
    case OBJ_PEN: {
      const width = view.getFloat32(base + 8, true);
      let color: string | undefined;
      for (let p = base + 12; p + 12 <= end; p += 4) {
        // The pen embeds its own solid-color brush block.
        if (view.getUint32(p, true) === GDIPLUS_VERSION && view.getUint32(p + 4, true) === 0) {
          color = argbHex(view.getUint32(p + 8, true));
          break;
        }
      }
      return { width, color };
    }
    case OBJ_PATH:
      return decodePath(payload, view, end, base);
    case OBJ_IMAGE:
      return decodeImage(payload, view, end, base);
    default:
      return undefined;
  }
}

/** GDI+ path persistence: pointCount, pointFormat, points, then one type
 *  byte per point. Type bits 0-2 carry the kind (start/line/bezier); bit 7
 *  marks the end of a closed figure. */
function decodePath(
  payload: Uint8Array,
  view: DataView,
  end: number,
  base: number,
): PathInfo | undefined {
  if (end < base + 12) return undefined;
  const count = view.getUint32(base, true);
  // Bit 14 (0x4000) switches PathPoints to int16 pairs. Other PathPointFlags
  // bits (e.g. the relative/RLE encoding the spec describes) are never emitted
  // by real GDI+ writers — corpus census: every non-zero-flag object decoded
  // by this single bit and nothing else.
  const compressed = (view.getUint32(base + 4, true) & 0x4000) !== 0;
  if (!count || count > 100_000) return undefined;
  const step = compressed ? 4 : 8;
  const pointsAt = base + 8;
  const typesAt = pointsAt + count * step;
  if (typesAt + count > end) return undefined;
  const cmds: PathCmds = [];
  let cx = 0,
    cy = 0;
  // A serialized path may open mid-figure (padded stream) — geometry before
  // the first start-type point has no anchor and would serialize as a `d`
  // beginning with L/C, which browsers drop wholesale.
  let started = false;
  let bezierTail: Array<[number, number]> = [];
  for (let i = 0; i < count; i++) {
    const p = pointsAt + i * step;
    cx = compressed ? view.getInt16(p, true) : view.getFloat32(p, true);
    cy = compressed ? view.getInt16(p + 2, true) : view.getFloat32(p + 4, true);
    const pt: [number, number] = [cx, cy];
    const t = payload[typesAt + i] & 0x07;
    if (t !== 0 && !started) continue;
    if (t === 3) {
      // A cubic consists of this point plus two more control points applied
      // to the running position.
      bezierTail.push(pt);
      if (bezierTail.length === 3) {
        cmds.push(["C", [...prevFlat(cmds), ...bezierTail.flat()]]);
        bezierTail = [];
      }
      continue;
    }
    if (t === 0) {
      started = true;
      cmds.push(["M", pt]);
    } else cmds.push(["L", pt]);
    if (payload[typesAt + i] & 0x80) cmds.push(["Z", []]);
  }
  return cmds.length >= 2 ? { cmds } : undefined;
}

/** The first coordinate pair of a pending cubic is the last plotted position
 *  (an implicit control anchor); failing that, zero — degenerate streams only. */
function prevFlat(cmds: PathCmds): number[] {
  for (let i = cmds.length - 1; i >= 0; i--) {
    const nums = cmds[i][1];
    if (nums.length >= 2) return [nums[nums.length - 2], nums[nums.length - 1]];
  }
  return [0, 0];
}

/** Natural pixel size of an embedded encoding, for normalizing DrawImagePoints'
 *  source rectangles into crop fractions: PNG IHDR / JPEG SOFn headers. */
function encodedSize(bytes: Uint8Array): { w: number; h: number } | undefined {
  // Signature bytes are byte-order-agnostic here; IHDR dimensions are BE.
  if (bytes.length >= 24 && view32(bytes, 0) === 0x89504e47) {
    return { w: view32(bytes, 16), h: view32(bytes, 20) };
  }
  if (bytes.length > 9 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    for (let p = 2; p + 9 < bytes.length;) {
      if (bytes[p] !== 0xff) break;
      const marker = bytes[p + 1];
      const len = (bytes[p + 2] << 8) | bytes[p + 3];
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return { w: (bytes[p + 7] << 8) | bytes[p + 8], h: (bytes[p + 5] << 8) | bytes[p + 6] };
      }
      p += 2 + len;
    }
  }
  return undefined;
}

function view32(b: Uint8Array, at: number): number {
  return new DataView(b.buffer, b.byteOffset + at, 4).getUint32(0, false);
}

interface ImageInfo {
  /** Bitmap encodings carry a data URL; metafile-typed images carry the raw
   *  nested EMF bytes instead (DrawImagePoints replays them as vectors). */
  src?: string;
  emfBytes?: Uint8Array;
  /** Natural pixel size of the encoding; drives DrawImagePoints src-rect crops. */
  w?: number;
  h?: number;
}

function decodeImage(
  payload: Uint8Array,
  view: DataView,
  end: number,
  base: number,
): ImageInfo | undefined {
  // ImageType word at `base`: 1 bitmap / 2 metafile. The metafile arm carries
  // a GDI+ metafile header followed by a complete EMF — replayed recursively
  // as vectors so text and shapes inside stay crisp at any scale.
  if (view.getUint32(base, true) === 2) {
    const start = nestedEmfStart(payload, end, base + 4);
    if (start >= 0) return { emfBytes: payload.subarray(start, end) };
    return undefined;
  }
  // Bitmap-typed images embed their original encoding; splice it straight
  // into a data URL instead of re-decoding pixels. The format's own end
  // marker bounds the slice — assembly prefixes and trailing run metadata
  // must not leak into the data URL.
  for (let p = base; p + 8 <= end; p++) {
    if (view.getUint32(p, true) === 0x474e5089 && view.getUint32(p + 4, true) === 0x0a1a0a0d) {
      let stop = end;
      for (let q = p; q + 8 <= end; q++) {
        if (view.getUint32(q, true) === 0x444e4549 && view.getUint32(q + 4, true) === 0x826042ae) {
          stop = q + 8;
          break;
        }
      }
      const bytes = payload.subarray(p, stop);
      const size = encodedSize(bytes);
      return { src: `data:image/png;base64,${base64(bytes)}`, ...size };
    }
  }
  for (let p = base; p + 3 <= end; p++) {
    if (payload[p] === 0xff && payload[p + 1] === 0xd8 && payload[p + 2] === 0xff) {
      let stop = -1;
      for (let q = end - 2; q >= p; q--) {
        if (payload[q] === 0xff && payload[q + 1] === 0xd9) {
          stop = q + 2;
          break;
        }
      }
      const bytes = payload.subarray(p, stop > 0 ? stop : end);
      const size = encodedSize(bytes);
      return { src: `data:image/jpeg;base64,${base64(bytes)}`, ...size };
    }
  }
  return undefined;
}

/** Offset of the nested EMF inside a metafile ImageData blob: after the GDI+
 *  header words, the stream opens with an EMR_HEADER whose record chain must
 *  validate through to EOF — candidate dwords that fail the walk are header
 *  fields that merely LOOK like type=1 records. */
function nestedEmfStart(payload: Uint8Array, end: number, from: number): number {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  for (let cand = from; cand + 108 <= end; cand += 4) {
    if (view.getUint32(cand, true) !== 1) continue;
    let off = cand;
    let steps = 0;
    while (off + 8 <= end && steps++ < 10_000) {
      const size = view.getUint32(off + 4, true);
      if (size < 8 || off + size > end) break;
      if (view.getUint32(off, true) === 14) return cand; // reached EOF
      off += size;
    }
  }
  return -1;
}

function base64(b: Uint8Array): string {
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < b.length; i += CHUNK) {
    bin += String.fromCharCode(...b.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

// ── replay ──

const MEMBERS_CAP = 4000;

/**
 * Replay a dual-mode WMF's embedded EMF+ stream into drawing members sized to
 * the display box. Returns undefined when there is nothing to replay: no
 * usable stream, no drawable output, or degenerate geometry.
 *
 * The Office exporter emits an object definition immediately before each call
 * that uses it, so draws take paint state from the most recently defined
 * brush/pen while the record's ObjectId selects the path being drawn (slot ids
 * pair up exactly that way across the corpus).
 */
export function emfPlusMembers(
  bytes: Uint8Array,
  boxW: number,
  boxH: number,
): LayoutDrawingMember[] | undefined {
  const emf = embeddedEmfStream(bytes);
  if (!emf) return undefined;
  const drafts = carrierDrafts(emf, undefined, 0);
  if (!drafts.length) return undefined;
  return finalize(drafts, boxW, boxH);
}

/** Nested-metafile recursion guard — a self-referencing or malformed blob
 *  must not spin through unbounded container levels. */
const MAX_NESTING = 3;

/** All drawable drafts of a raw carrier EMF: its concatenated EMF+ comment
 *  stream replayed with `basis` pre-applied on top of every world transform,
 *  plus the GDI-side text chain. `basis` carries nested DrawImagePoints
 *  parallelograms: the parent maps the image rect onto it and the child plays
 *  under that mapping. */
function carrierDrafts(emf: Uint8Array, basis: Xform | undefined, depth: number): Draft[] {
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
  // Every draft coordinate passes through this effective transform first; a
  // nested level folds its parent's image-to-parallelogram basis here.
  const effOf = (xf: Xform): Xform => (basis ? combine(basis, xf) : xf);

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
        if (path?.cmds && lastBrush?.solid)
          pushPath(drafts, path.cmds, effOf(xf), { fill: lastBrush.solid });
        break;
      }
      case PLUS_DRAW_PATH: {
        installOpenRuns();
        const path = objects.get(objectId) as PathInfo | undefined;
        if (path?.cmds && lastPen?.color) {
          pushPath(drafts, path.cmds, effOf(xf), {
            strokeColor: lastPen.color,
            strokeWidth: lastPen.width,
          });
        }
        break;
      }
      case PLUS_FILL_RECTS: {
        installOpenRuns();
        const n = pv.getUint32(d, true);
        if (!n || n > 10_000 || d + 4 + n * 16 > d + rs - 8) break;
        const eff = effOf(xf);
        for (let i = 0; i < n; i++) {
          const rx = pv.getFloat32(d + 4 + i * 16, true);
          const ry = pv.getFloat32(d + 8 + i * 16, true);
          const rw = pv.getFloat32(d + 12 + i * 16, true);
          const rh = pv.getFloat32(d + 16 + i * 16, true);
          const [x, y] = xformPoint(eff, rx, ry);
          const [x2, y2] = xformPoint(eff, rx + rw, ry + rh);
          pushRect(drafts, x, y, x2 - x, y2 - y, lastBrush?.solid);
        }
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
  // Dual-mode carriers keep the real text on the GDI side; merge those runs
  // in before deciding whether anything rendered at all.
  drafts.push(...gdiTextDrafts(emf, effOf));
  return drafts;
}

/** One metafile-typed EmfPlusImage drawn onto a destination parallelogram:
 *  map the nested EMF's device rectangle onto the mapped corners, then play
 *  the nested carrier with that mapping as its basis. */
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
  const end = Math.min(view.byteLength, d + rs);
  const pts = d + 32;
  if (pts + 12 > end || view.getUint32(d + 28, true) !== 3) return;
  const compressed = (flags & 0x4000) !== 0;
  const read = (i: number): [number, number] => {
    const p = pts + i * (compressed ? 4 : 8);
    return compressed
      ? [view.getInt16(p, true), view.getInt16(p + 2, true)]
      : [view.getFloat32(p, true), view.getFloat32(p + 4, true)];
  };
  const corners = [read(0), read(1), read(2)].map(([x, y]) => xformPoint(outerEff, x, y));
  const b = emfDeviceRect(nested);
  if (!b || b.w <= 0 || b.h <= 0) return;
  const [p0, p1, p2] = corners;
  // Affine taking nested device coords onto the drawn parallelogram.
  const basis: Xform = {
    m11: (p1[0] - p0[0]) / b.w,
    m21: (p1[1] - p0[1]) / b.w,
    m12: (p2[0] - p0[0]) / b.h,
    m22: (p2[1] - p0[1]) / b.h,
    dx: p0[0],
    dy: p0[1],
  };
  drafts.push(...carrierDrafts(nested, basis, depth + 1));
}

/** rclBounds of a raw EMF's first EMR_HEADER ([+8..24] RECT ints). */
function emfDeviceRect(emf: Uint8Array): { w: number; h: number } | undefined {
  if (emf.length < 26) return undefined;
  const v = new DataView(emf.buffer, emf.byteOffset, emf.byteLength);
  const l = v.getInt32(8, true);
  const t = v.getInt32(12, true);
  const r = v.getInt32(16, true);
  const bo = v.getInt32(20, true);
  return { w: r - l, h: bo - t };
}

function pushPath(
  drafts: Draft[],
  rawCmds: PathCmds,
  xf: Xform,
  paint: { fill?: string; strokeColor?: string; strokeWidth?: number },
): void {
  const transformed: PathCmds = rawCmds.map(([op, nums]) => {
    if (op === "Z") return [op, nums];
    const out: number[] = [];
    for (let i = 0; i + 1 < nums.length; i += 2) {
      const [x, y] = xformPoint(xf, nums[i], nums[i + 1]);
      out.push(x, y);
    }
    return [op, out];
  });
  drafts.push({
    kind: "path",
    cmds: transformed,
    ...(paint.fill ? { fill: paint.fill } : {}),
    ...(paint.strokeColor && paint.strokeWidth != null
      ? { strokeColor: paint.strokeColor, strokeWidth: paint.strokeWidth }
      : {}),
  });
}

function pushRect(
  drafts: Draft[],
  x: number,
  y: number,
  w: number,
  h: number,
  fill?: string,
): void {
  if (!fill || !(w >= 0 && h >= 0)) return;
  drafts.push({
    kind: "path",
    cmds: [
      ["M", [x, y]],
      ["L", [x + w, y]],
      ["L", [x + w, y + h]],
      ["L", [x, y + h]],
      ["Z", []],
    ],
    fill,
  });
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
  const end = Math.min(buf.length, d + rs);
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

// ── GDI-side text ──

// Carrier EMR codes consumed while replaying the text chain ([MS-EMF]
// RecordType values).
const EMR_SAVEDC = 33;
const EMR_RESTOREDC = 34;
const EMR_SET_WORLD_TRANSFORM = 35;
const EMR_MODIFY_WORLD_TRANSFORM = 36;
const EMR_SELECT_OBJECT = 37;
const EMR_DELETE_OBJECT = 40;
const EMR_SETTEXTCOLOR = 24;
const EMR_EXT_CREATE_FONT = 82;
const EMR_EXT_TEXT_OUT_A = 83;
const EMR_EXT_TEXT_OUT_W = 84;

/** EmrText offsets inside EMR_EXTTEXTOUTW (relative to the record start):
 *  bounds[16], graphicsMode, ex/eyScale, then Reference xy, Chars, offString,
 *  Options (+ an optional rect). The string lives at offString, relative to
 *  the record start. */
const EXT_REF_X = 36;
const EXT_REF_Y = 40;
const EXT_CHARS = 44;
const EXT_OFF_STRING = 48;

interface GdiFont {
  height: number;
  weight: number;
  face?: string;
}

/** Raw EMR-world XFORM: six floats directly behind the record header (the
 *  byte-length prefix seen on EMF+ payloads does not appear here). */
function readCarrierXform(view: DataView, at: number): Xform {
  return {
    m11: view.getFloat32(at, true),
    m12: view.getFloat32(at + 4, true),
    m21: view.getFloat32(at + 8, true),
    m22: view.getFloat32(at + 12, true),
    dx: view.getFloat32(at + 16, true),
    dy: view.getFloat32(at + 20, true),
  };
}

/** Replay only the text chain of the carrier's GDI records and project it into
 *  the same display space the EMF+ drafts occupy. Dual-mode files keep their
 *  real text as ExtTextOutW runs drawn in arbitrary logical scales; each run's
 *  placement comes from the carrier's live world transform, which folds every
 *  domain onto the EMF's device rectangle (rclBounds) shared with the EMF+
 *  output. Fonts are object-table based (CreateFontIndirectW defines a slot,
 *  SelectObject activates it); color rides on SetTextColor's COLORREF.
 *  `effOf` folds a parent nesting basis into that transform (nested carriers). */
function gdiTextDrafts(emf: Uint8Array, effOf?: (xf: Xform) => Xform): TextDraft[] {
  const view = new DataView(emf.buffer, emf.byteOffset);
  const drafts: TextDraft[] = [];
  const fonts = new Map<number, GdiFont>();
  let selected = -1;
  let xf: Xform = { ...IDENTITY };
  let textColor: string | undefined;
  // SaveDC snapshots the full drawing state we track.
  const states: Array<{ xf: Xform; selected: number; textColor?: string }> = [];
  for (let eo = 0; eo + 8 <= emf.length;) {
    const type = view.getUint32(eo, true);
    const size = view.getUint32(eo + 4, true);
    if (size < 8 || eo + size > emf.length) break;
    switch (type) {
      case EMR_EXT_CREATE_FONT: {
        // EXTCREATEFONTINDIRECTW: object index dword + LOGFONTW whose face
        // name is UTF-16LE at byte 28 of the LOGFONT.
        const slot = view.getUint32(eo + 8, true);
        const base = eo + 12;
        let face = "";
        for (let c = 0; c < 32 && base + 28 + c * 2 + 1 < eo + size; c++) {
          const ch = view.getUint16(base + 28 + c * 2, true);
          if (!ch) break;
          face += String.fromCharCode(ch);
        }
        fonts.set(slot, {
          height: view.getInt32(base, true),
          weight: view.getUint32(base + 16, true),
          face,
        });
        break;
      }
      case EMR_SELECT_OBJECT: {
        // Stock objects carry a flag bit; real slots reference our table.
        const raw = view.getUint32(eo + 8, true);
        if (!(raw & 0x80000000) && fonts.has(raw)) selected = raw;
        break;
      }
      case EMR_DELETE_OBJECT:
        fonts.delete(view.getUint32(eo + 8, true));
        break;
      case EMR_SETTEXTCOLOR: {
        // COLORREF (0x00bbggrr); high-flag words reference the system palette
        // — skip those and keep the last true RGB.
        const ref = view.getUint32(eo + 8, true);
        if (ref <= 0xffffff) textColor = rgbHex(ref);
        break;
      }
      case EMR_SAVEDC:
        states.push({ xf, selected, textColor });
        break;
      case EMR_RESTOREDC: {
        const st = states.pop();
        if (st) {
          ({ xf, selected, textColor } = st);
        }
        break;
      }
      case EMR_SET_WORLD_TRANSFORM:
        xf = readCarrierXform(view, eo + 8);
        break;
      case EMR_MODIFY_WORLD_TRANSFORM: {
        const next = readCarrierXform(view, eo + 8);
        const mode = view.getUint32(eo + 32, true);
        // ModifyWorldTransformMode: 1 resets to identity, 2/3 multiply naming
        // the operand side, 4 (MWT_SET — the corpus norm) overwrites outright.
        xf =
          mode === 1
            ? { ...IDENTITY }
            : mode === 2
              ? combine(next, xf)
              : mode === 3
                ? combine(xf, next)
                : next;
        break;
      }
      case EMR_EXT_TEXT_OUT_W:
      case EMR_EXT_TEXT_OUT_A: {
        // UTF-16 / one-byte-per-char strings; only W appears in the corpus,
        // both decode cheaply.
        const chars = view.getUint32(eo + EXT_CHARS, true);
        const offString = view.getUint32(eo + EXT_OFF_STRING, true);
        if (!chars || chars > 4096) break;
        const strAt = eo + offString;
        const strEnd = strAt + chars * (type === EMR_EXT_TEXT_OUT_W ? 2 : 1);
        if (strEnd > eo + size) break;
        let text = "";
        if (type === EMR_EXT_TEXT_OUT_W) {
          for (let c = 0; c < chars; c++)
            text += String.fromCharCode(view.getUint16(strAt + c * 2, true));
        } else {
          const gbk = new TextDecoder("gbk");
          text = gbk.decode(emf.subarray(strAt, strEnd));
        }
        // The walk loop has no update clause — never `continue` past the
        // switch below or a single record spins forever.
        if (text.trim()) {
          const font = fonts.get(selected);
          // Logical draw origin → carrier device space through the live world
          // transform (folded with any nesting basis); glyph metrics scale with
          // its uniform factor.
          const eff = effOf ? effOf(xf) : xf;
          const x = view.getInt32(eo + EXT_REF_X, true);
          const yBaseline = view.getInt32(eo + EXT_REF_Y, true);
          const height =
            Math.abs(font?.height ?? 100) * ((Math.abs(eff.m11) + Math.abs(eff.m22)) / 2);
          // Baseline origin → box top by a typical CJK ascent (same calibration
          // as the WMF player); advance falls back to per-char estimates because
          // the trail Dx run is not worth parsing for outline-position accuracy.
          let advance = 0;
          for (const ch of text) advance += height * (ch.charCodeAt(0) > 0xff ? 1 : 0.55);
          drafts.push({
            kind: "text",
            x: x * eff.m11 + yBaseline * eff.m21 + eff.dx,
            y: x * eff.m12 + yBaseline * eff.m22 + eff.dy - height * 0.8,
            w: advance + 2,
            h: height * 1.35,
            text,
            family: font?.face ?? "",
            sizeWorld: height,
            ...(textColor ? { color: textColor } : {}),
            ...(font && font.weight >= 550 ? { bold: true } : {}),
          });
        }
        break;
      }
      default:
        break;
    }
    eo += size;
  }
  return drafts;
}

function rgbHex(colorref: number): string {
  const r = colorref & 0xff;
  const g = (colorref >> 8) & 0xff;
  const b = (colorref >> 16) & 0xff;
  return `${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0")}`;
}

/** Whether a draft joins the union bounding box that drives the final
 *  normalization. After carrier world-transform projection all drafts share
 *  one device space, so text participates like any other content.
 *  Zero-area strokes join nothing: they render as hairlines at best yet can
 *  drag the box to a far corner. */
function participatesInBox(dr: Draft): boolean {
  if (dr.kind === "path") {
    const b = boxOf(dr);
    return b.x1 - b.x0 > 1 && b.y1 - b.y0 > 1;
  }
  return true;
}

/** Scale drafts from EMF device coordinates into the display box and shape
 *  the renderer-facing members. Independent axis scales preserve relative
 *  layout; label proportions ride along because text arrives as outlines. */
function finalize(drafts: Draft[], boxW: number, boxH: number): LayoutDrawingMember[] | undefined {
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity;
  for (const dr of drafts) {
    if (!participatesInBox(dr)) continue;
    if (dr.kind === "path") {
      for (const [, nums] of dr.cmds) {
        for (let i = 0; i + 1 < nums.length; i += 2) {
          minX = Math.min(minX, nums[i]);
          maxX = Math.max(maxX, nums[i]);
          minY = Math.min(minY, nums[i + 1]);
          maxY = Math.max(maxY, nums[i + 1]);
        }
      }
      continue;
    }
    minX = Math.min(minX, dr.x);
    minY = Math.min(minY, dr.y);
    maxX = Math.max(maxX, dr.x + dr.w);
    maxY = Math.max(maxY, dr.y + dr.h);
  }
  const bw = maxX - minX,
    bh = maxY - minY;
  if (!(bw > 0 && bh > 0)) return undefined;
  const sX = boxW / bw,
    sY = boxH / bh;
  const X = (x: number) => (x - minX) * sX;
  const Y = (y: number) => (y - minY) * sY;
  const members: LayoutDrawingMember[] = [];
  for (const dr of drafts) {
    if (dr.kind === "text") {
      members.push({
        kind: "textBox",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        blocks: [
          {
            kind: "paragraph",
            inline: [
              {
                kind: "text",
                text: dr.text,
                style: {
                  family: dr.family,
                  sizePx: dr.sizeWorld * sY,
                  ...(dr.color ? { color: dr.color } : {}),
                  ...(dr.bold ? { bold: true } : {}),
                },
              },
            ],
          },
        ],
      });
      continue;
    }
    if (dr.kind === "pic") {
      members.push({
        kind: "picture",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        src: dr.src,
        ...(dr.crop
          ? {
              crop: {
                left: Math.max(0, Math.min(1, dr.crop.l)),
                top: Math.max(0, Math.min(1, dr.crop.t)),
                right: Math.max(0, Math.min(1, dr.crop.r)),
                bottom: Math.max(0, Math.min(1, dr.crop.b)),
              },
            }
          : {}),
      });
      continue;
    }
    const parts: string[] = [];
    for (const [op, nums] of dr.cmds) {
      if (op === "Z") {
        parts.push("Z");
        continue;
      }
      const coords: string[] = [];
      for (let i = 0; i + 1 < nums.length; i += 2) {
        coords.push(round1(X(nums[i])), round1(Y(nums[i + 1])));
      }
      parts.push(op + coords.join(" "));
    }
    if (parts.length < 2) continue;
    const strokeWidth =
      dr.strokeWidth != null
        ? Math.max(1, Math.round(dr.strokeWidth * ((sX + sY) / 2)))
        : undefined;
    members.push({
      kind: "path",
      x: 0,
      y: 0,
      width: boxW,
      height: boxH,
      d: parts.join(""),
      ...(dr.fill ? { fill: dr.fill, fillRule: "evenodd" as const } : {}),
      ...(dr.strokeColor && strokeWidth
        ? { line: { px: strokeWidth, color: dr.strokeColor } }
        : {}),
    });
    if (members.length > MEMBERS_CAP) return undefined;
  }
  return members.length ? members : undefined;
}

function round1(n: number): string {
  return String(Math.round(n * 10) / 10);
}

/** World-space bounding box of one draft (diagnostics hook shape). */
function boxOf(dr: Draft): { x0: number; y0: number; x1: number; y1: number } {
  if (dr.kind !== "path") {
    return { x0: dr.x, y0: dr.y, x1: dr.x + dr.w, y1: dr.y + dr.h };
  }
  let x0 = Infinity,
    y0 = Infinity,
    x1 = -Infinity,
    y1 = -Infinity;
  for (const [, nums] of dr.cmds) {
    for (let i = 0; i + 1 < nums.length; i += 2) {
      x0 = Math.min(x0, nums[i]);
      x1 = Math.max(x1, nums[i]);
      y0 = Math.min(y0, nums[i + 1]);
      y1 = Math.max(y1, nums[i + 1]);
    }
  }
  return { x0, y0, x1, y1 };
}
