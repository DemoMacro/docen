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
}

type Draft = PathDraft | PicDraft;

// ── object decoding ──
// EmfPlusObject carries TotalObjectSize at payload [+8], the shared version
// stamp at [+12], and per-type fields from [+16].

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
interface ImageInfo {
  src: string;
}

function decodeObject(
  view: DataView,
  po: number,
  rs: number,
  buf: Uint8Array,
): BrushInfo | PenInfo | PathInfo | ImageInfo | undefined {
  const type = (view.getUint16(po + 2, true) >> 8) & 0x7f;
  const end = po + rs;
  if (po + 24 > end) return undefined;
  switch (type) {
    case OBJ_BRUSH: {
      const brushType = view.getUint32(po + 16, true);
      if (brushType === 0) return { solid: argbHex(view.getUint32(po + 20, true)) };
      // PathGradient brushes take their declared center color as the flat
      // approximation — the member protocol paints solids only, so a bounded
      // scan for the first bright opaque word stands in until gradient fills
      // become a member-level concept.
      for (let p = po + 20; p < Math.min(end - 4, po + 160); p += 4) {
        const c = view.getUint32(p, true);
        if (c >>> 24 !== 0xff) continue;
        const rgb = c & 0xffffff;
        const lum = ((rgb >> 16) & 0xff) * 299 + ((rgb >> 8) & 0xff) * 587 + (rgb & 0xff) * 114;
        if (lum > 120_000) return { solid: argbHex(c) };
      }
      return undefined;
    }
    case OBJ_PEN: {
      const width = view.getFloat32(po + 24, true);
      let color: string | undefined;
      for (let p = po + 28; p + 12 <= end; p += 4) {
        // The pen embeds its own solid-color brush block.
        if (view.getUint32(p, true) === GDIPLUS_VERSION && view.getUint32(p + 4, true) === 0) {
          color = argbHex(view.getUint32(p + 8, true));
          break;
        }
      }
      return { width, color };
    }
    case OBJ_PATH:
      return decodePath(view, po, end, buf);
    case OBJ_IMAGE:
      return decodeImage(view, po, end, buf);
    default:
      return undefined;
  }
}

/** GDI+ path persistence: pointCount, pointFormat, points, then one type
 *  byte per point. Type bits 0-2 carry the kind (start/line/bezier); bit 7
 *  marks the end of a closed figure. */
function decodePath(
  view: DataView,
  po: number,
  end: number,
  buf: Uint8Array,
): PathInfo | undefined {
  if (po + 28 > end) return undefined;
  const count = view.getUint32(po + 16, true);
  const compressed = (view.getUint32(po + 20, true) & 0x2000) !== 0;
  if (!count || count > 100_000) return undefined;
  const step = compressed ? 4 : 8;
  const typesAt = po + 24 + count * step;
  if (typesAt + count > end) return undefined;
  const cmds: PathCmds = [];
  let cx = 0,
    cy = 0;
  let bezierTail: Array<[number, number]> = [];
  for (let i = 0; i < count; i++) {
    const p = po + 24 + i * step;
    cx = compressed ? view.getInt16(p, true) : view.getFloat32(p, true);
    cy = compressed ? view.getInt16(p + 2, true) : view.getFloat32(p + 4, true);
    const pt: [number, number] = [cx, cy];
    const t = buf[typesAt + i] & 0x07;
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
    if (t === 0) cmds.push(["M", pt]);
    else cmds.push(["L", pt]);
    if (buf[typesAt + i] & 0x80) cmds.push(["Z", []]);
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

function decodeImage(
  view: DataView,
  po: number,
  end: number,
  buf: Uint8Array,
): ImageInfo | undefined {
  // Bitmap-typed images embed their original encoding; splice it straight
  // into a data URL instead of re-decoding pixels.
  for (let p = po + 20; p + 8 <= end; p++) {
    if (view.getUint32(p, true) === 0x474e5089 && view.getUint32(p + 4, true) === 0x0a1a0a0d) {
      return { src: `data:image/png;base64,${base64(buf.subarray(p, end))}` };
    }
  }
  for (let p = po + 20; p + 3 <= end; p++) {
    if (buf[p] === 0xff && buf[p + 1] === 0xd8 && buf[p + 2] === 0xff) {
      return { src: `data:image/jpeg;base64,${base64(buf.subarray(p, end))}` };
    }
  }
  return undefined;
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

  // Concatenate every EMR_GDICOMMENT "EMF+" payload into one record stream.
  const ev = new DataView(emf.buffer);
  let plusLen = 0;
  for (let eo = 0; eo + 8 <= emf.length;) {
    const type = ev.getUint32(eo, true);
    const size = ev.getUint32(eo + 4, true);
    if (size < 8 || eo + size > emf.length) break;
    if (type === 70 && ev.getUint32(eo + 12, true) === 0x2b464d45) plusLen += size - 16;
    eo += size;
  }
  if (!plusLen) return undefined;
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

  const pv = new DataView(plus.buffer);
  const objects = new Map<number, unknown>();
  let xf: Xform = { ...IDENTITY };
  const saved: Xform[] = [];
  const drafts: Draft[] = [];
  let lastBrush: BrushInfo | undefined;
  let lastPen: PenInfo | undefined;

  for (let po = 0; po + 8 <= plus.length;) {
    const rt = pv.getUint16(po, true);
    const flags = pv.getUint16(po + 2, true);
    const rs = pv.getUint32(po + 4, true);
    if (rs < 8 || po + rs > plus.length) break;
    if (rt === PLUS_END_OF_FILE) break;
    const d = po + 8;
    const objectId = flags & 0xff;
    switch (rt) {
      case PLUS_OBJECT: {
        const obj = decodeObject(pv, po, rs, plus);
        if (!obj) break;
        if ("solid" in obj) lastBrush = obj;
        else if ("width" in obj) lastPen = obj;
        objects.set(objectId, obj);
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
        const path = objects.get(objectId) as PathInfo | undefined;
        if (path?.cmds && lastBrush?.solid)
          pushPath(drafts, path.cmds, xf, { fill: lastBrush.solid });
        break;
      }
      case PLUS_DRAW_PATH: {
        const path = objects.get(objectId) as PathInfo | undefined;
        if (path?.cmds && lastPen?.color) {
          pushPath(drafts, path.cmds, xf, {
            strokeColor: lastPen.color,
            strokeWidth: lastPen.width,
          });
        }
        break;
      }
      case PLUS_FILL_RECTS: {
        const n = pv.getUint32(d, true);
        if (!n || n > 10_000 || d + 4 + n * 16 > d + rs - 8) break;
        for (let i = 0; i < n; i++) {
          const rx = pv.getFloat32(d + 4 + i * 16, true);
          const ry = pv.getFloat32(d + 8 + i * 16, true);
          const rw = pv.getFloat32(d + 12 + i * 16, true);
          const rh = pv.getFloat32(d + 16 + i * 16, true);
          const [x, y] = xformPoint(xf, rx, ry);
          const [x2, y2] = xformPoint(xf, rx + rw, ry + rh);
          pushRect(drafts, x, y, x2 - x, y2 - y, lastBrush?.solid);
        }
        break;
      }
      case PLUS_DRAW_IMAGE_POINTS: {
        const img = objects.get(objectId) as ImageInfo | undefined;
        if (img?.src) drawImagePoints(pv, plus, d, rs, flags, xf, img.src, drafts);
        break;
      }
      default:
        break;
    }
    po += rs;
  }
  if (!drafts.length) return undefined;
  return finalize(drafts, boxW, boxH);
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
  src: string,
  drafts: Draft[],
): void {
  // Destination parallelogram: after a 28-byte metadata head, a word count
  // followed by three points — compressed int16 pairs under the P flag,
  // float32 pairs otherwise.
  const end = Math.min(buf.length, d + rs);
  const pts = d + 32;
  if (pts + 12 > end || view.getUint32(d + 28, true) !== 3) return;
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
  drafts.push({ kind: "pic", x: minX, y: minY, w, h, src });
}

/** Scale drafts from EMF world coordinates into the display box and shape the
 *  renderer-facing members. Independent axis scales preserve relative layout;
 *  label proportions ride along because text arrives as outlines. */
function finalize(drafts: Draft[], boxW: number, boxH: number): LayoutDrawingMember[] | undefined {
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity;
  for (const dr of drafts) {
    if (dr.kind === "pic") {
      minX = Math.min(minX, dr.x);
      minY = Math.min(minY, dr.y);
      maxX = Math.max(maxX, dr.x + dr.w);
      maxY = Math.max(maxY, dr.y + dr.h);
      continue;
    }
    for (const [, nums] of dr.cmds) {
      for (let i = 0; i + 1 < nums.length; i += 2) {
        minX = Math.min(minX, nums[i]);
        maxX = Math.max(maxX, nums[i]);
        minY = Math.min(minY, nums[i + 1]);
        maxY = Math.max(maxY, nums[i + 1]);
      }
    }
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
    if (dr.kind === "pic") {
      members.push({
        kind: "picture",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        src: dr.src,
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
