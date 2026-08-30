// Office renders clipboard/import vector media as dual-mode metafiles: the
// placeable-WMF body carries only a coarse GDI approximation, while
// META_ESCAPE chunks (magic "WMFC") embed the real drawing as a complete EMF
// whose EMR_GDICOMMENT payloads hold the GDI+ (EMF+) record stream. This
// module reassembles that stream and replays it into the same structured
// drawing members native OOXML drawings are projected into, so metafile art
// renders through the identical layout/paint model instead of a raster detour.

import { bmpDataUrl } from "./dib";
import type { MetafileMember, SourceCrop } from "./member";

export type { SourceCrop } from "./member";

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
const PLUS_DRAW_LINES = 0x400d;
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

/** ARGB word (0xAARRGGBB) → hex RRGGBB composited over the white page;
 *  undefined when fully transparent. Metafiles wash whole-page tints at
 *  near-zero alpha (a 10/255 red wash rendered opaque covers the art under
 *  it), so partial alpha pre-blends with white instead of riding along. */
function argbHex(word: number): string | undefined {
  const a = word >>> 24;
  if (a === 0) return undefined;
  const rgb = word & 0xffffff;
  if (a >= 0xfa) return (rgb | 0x1000000).toString(16).slice(1);
  const mix = (c: number) => Math.round((c * a + 255 * (255 - a)) / 255);
  return [mix(rgb >> 16), mix((rgb >> 8) & 0xff), mix(rgb & 0xff)]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("");
}

// A sub-path as raw command tuples; "M" starts, "L"/"C" continue, "Z" closes.
type PathCmds = Array<["M" | "L" | "C" | "Z", number[]]>;

interface PathDraft {
  kind: "path";
  cmds: PathCmds;
  fill?: string;
  strokeWidth?: number;
  strokeColor?: string;
  /** Preset dash token — threaded to the member's line.dash verbatim. */
  dash?: string;
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
  /** Ternary raster-op emulation blend (SRCPAINT → screen, SRCAND → multiply). */
  blend?: "screen" | "multiply";
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
  /** Extra per-character advance from the trail Dx run, device px — GDI
   *  tracked text (letter-spaced labels) reports wider advances than the
   *  glyphs are wide; the member threads the difference into the run style. */
  letterSpacingWorld?: number;
  /** Cross-axis inset of the ink inside its cell, device px — a vertical
   *  (@font) run centers each glyph in its cell, so the ink starts
   *  (cellW − em)/2 into the box the fold produced. */
  cellInsetWorld?: number;
  /** Clockwise screen angle of the text-direction column, degrees — a
   *  rotated world transform means vertical text (plan-box columns); the
   *  member carries it so the paint rotates instead of laying the run
   *  horizontally across neighboring columns. */
  rotation?: number;
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
  /** GDI+ DashStyle preset token (a prstDash value) — the pen's line style. */
  dash?: string;
}

/** GDI+ DashStyle enum values → prstDash tokens (the member protocol's dash
 *  vocabulary); Solid stays absent. */
const DASH_STYLE_TOKENS = ["", "", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"];

/** Custom DashStyle (5) carries a repeat array in pen-width units; the corpus
 *  emits exactly the two stock shapes ([3,1] Dash, [1,1] Dot), so only those
 *  map to tokens — anything else falls back to solid rather than inventing a
 *  pattern the renderer has no verified sample for. */
function dashTokenFor(dashes: number[]): string | undefined {
  const key = dashes.map((n) => Math.round(n)).join(",");
  if (key === "3,1") return "sysDash";
  if (key === "1,1") return "sysDot";
  return undefined;
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
      // EmfPlusPen layout (corpus-verified): BrushId u32, PenDataFlags u32,
      // Unit u32, Width REAL — the trailing brush blob carries the color.
      const flags = view.getUint32(base + 4, true);
      const width = view.getFloat32(base + 12, true);
      let cursor = base + 16;
      // OptionalData fields consumed in ascending PenDataFlags bit order
      // ([MS-EMFPLUS] §2.2.2.28): transform(6 REALs), start/end cap, join,
      // miter limit, dash style DWORD, dashed-line cap, dashed-line offset
      // REAL, then the length-prefixed dash array; caps/compound arrays and
      // non-center/alignment words beyond that are skipped unparsed.
      if (flags & 0x0001) cursor += 24;
      if (flags & 0x0002) cursor += 4;
      if (flags & 0x0004) cursor += 4;
      if (flags & 0x0008) cursor += 4;
      if (flags & 0x0010) cursor += 4;
      let style: number | undefined;
      if (flags & 0x0020) {
        style = view.getUint32(cursor, true);
        cursor += 4;
      }
      if (flags & 0x0040) cursor += 4;
      if (flags & 0x0080) cursor += 4;
      let dashes: number[] | undefined;
      if (flags & 0x0100) {
        const n = view.getUint32(cursor, true);
        if (n > 0 && n <= 16 && cursor + 4 + n * 4 <= end) {
          dashes = [];
          for (let i = 0; i < n; i++) dashes.push(view.getFloat32(cursor + 4 + i * 4, true));
        }
        cursor += 4 + Math.min(n ?? 0, 16) * 4;
      }
      if (flags & 0x0200) cursor += 4;
      if (flags & 0x0400) {
        const n = view.getUint32(cursor, true);
        cursor += 4 + Math.min(n, 64) * 4;
      }
      const dash =
        dashes != null ? dashTokenFor(dashes) : DASH_STYLE_TOKENS[style ?? 0] || undefined;
      let color: string | undefined;
      // The pen's own brush follows all optional fields ([totalSize][version]
      // stamped), then one byte-size word + a zero word before the ARGB.
      for (let p = cursor; p + 12 <= end; p += 4) {
        if ((view.getUint32(p, true) & 0xffff0000) === (GDIPLUS_VERSION & 0xffff0000)) {
          color = argbHex(view.getUint32(p + 8, true));
          break;
        }
      }
      if (!color) {
        // Gradient-brush pens carry no solid stamp in the scanned shape: their
        // blob is a LinearGradient whose visible color lives in the preset
        // ramp behind the [flags][wrap][rect] prefix — alpha ramps from fully
        // transparent up to the opaque end that dominates the rendered stroke
        // (its interior 0xff000000 words are header fields, not the color).
        // Take the last fully-opaque word as the flat approximation; the
        // strict alpha test keeps the 0xdb-alpha version stamp out.
        for (let p = end - 4; p >= cursor; p -= 4) {
          if (view.getUint32(p, true) >>> 24 === 0xff) {
            color = argbHex(view.getUint32(p, true));
            break;
          }
        }
      }
      return { width, ...(dash ? { dash } : {}), color };
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
      // A GDI+ bezier triplet is [control1, control2, end] — the segment opens
      // at the running position, so the serialized cubic is exactly these
      // three pairs (an SVG "C" with a leading anchor would be malformed).
      bezierTail.push(pt);
      if (bezierTail.length === 3) {
        cmds.push(["C", bezierTail.flat()]);
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

/** CJK cell-height / em ratio for positive GDI lfHeight values (SimSun's OS/2
 *  winAscent+winDescent). Shared with the WMF player's font sizing. */
export const GDI_CELL_PER_EM = 1.2969;

/** Realized GDI ascent/descent per em for a vertical (@-prefixed) CJK font —
 *  measured on @楷体 at a −360 request (tmAscent 328 / tmDescent 60), the
 *  corpus's only vertical-slogan face. Word's metafile replay renders these
 *  runs upright with the glyph BASELINE through the reference point, so the
 *  ink centers (ascent − descent)/2 = 0.37 em right of it. */
const GDI_V_ASCENT_PER_EM = 328 / 360;
const GDI_V_DESCENT_PER_EM = 60 / 360;

/** CJK punctuation that Word's vertical replay turns 90° clockwise (the
 *  font's vert substitutions): brackets, quotes, dots and dashes. Ideographs,
 *  kana, fullwidth alphanumerics, 々/〇 and the katakana middle dot stay
 *  upright — their vert forms are identity. */
function isVerticalPunct(ch: string): boolean {
  const c = ch.codePointAt(0) ?? 0;
  if (c === 0x3005 || c === 0x3007 || c === 0x30fb) return false;
  if (c >= 0x3000 && c <= 0x303f) return true; // 、。〈〉《》「」『』【】〔〕…
  if (c >= 0x2013 && c <= 0x2026) return true; // – — ‘ ’ “ ” ‥ …
  if (c >= 0xff01 && c <= 0xff0f) return true; // ！＂＃…（）＊＋，－．／
  if (c >= 0xff1a && c <= 0xff20) return true; // ：；＜＝＞？＠
  if (c >= 0xff3b && c <= 0xff40) return true; // ［＼］＾＿｀
  if (c >= 0xff5b && c <= 0xff65) return true; // ｛｜｝～ and halfwidth 、｡｢｣
  return false;
}

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
  crop?: SourceCrop,
): MetafileMember[] | undefined {
  const emf = embeddedEmfStream(bytes);
  if (!emf) return undefined;
  const drafts = carrierDrafts(emf, undefined, 0);
  if (!drafts.length) return undefined;
  // The carrier's physical frame (EMR_HEADER rclFrame, [MS-EMF] §2.3.4.2)
  // anchors the display mapping, converted to EMF+ device px at 96 dpi — the
  // reference resolution these clipboard metafiles declare their extent at.
  // rclBounds is only an ink hint in a device basis that synthetic carriers
  // leave inconsistent with the frame; stretching bounds onto the extent
  // distorts the art Word renders at the frame's near-1:1 scale.
  let frame: { x: number; y: number; w: number; h: number } | undefined;
  if (emf.length >= 40) {
    const v = new DataView(emf.buffer, emf.byteOffset, emf.byteLength);
    if (v.getUint32(0, true) === 1 && v.getUint32(4, true) >= 88) {
      const l = (v.getInt32(24, true) * 96) / 2540,
        t = (v.getInt32(28, true) * 96) / 2540,
        r = (v.getInt32(32, true) * 96) / 2540,
        b = (v.getInt32(36, true) * 96) / 2540;
      if (r > l && b > t) frame = { x: l, y: t, w: r - l, h: b - t };
    }
  }
  return finalize(drafts, boxW, boxH, frame, crop);
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
        const n = pv.getUint32(d + 8, true);
        const step = flags & 0x4000 ? 8 : 16;
        if (!n || n > 10_000 || d + 12 + n * step > d + rs) break;
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
        const n = pv.getUint32(d + 4, true);
        const compressed = (flags & 0x8000) !== 0;
        const step = compressed ? 4 : 8;
        if (!n || n > 10_000 || d + 8 + n * step > d + rs) break;
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
  // (a bare EMF+ header with all imagery on the GDI side). GDI paths keep
  // the coincident-box dedup only: unlike blts, corpus files do split their
  // vector art across both layers, so a blanket drop would delete art. Text
  // never drops: the EMF+ walk does not decode DrawString runs yet, and a
  // dual file's GDI text is the same glyphs at the same spots.
  const plusPaints = drafts
    .filter((dr): dr is PathDraft => dr.kind === "path" && (!!dr.fill || !!dr.strokeColor))
    .map((dr) => ({ fill: dr.fill, stroke: dr.strokeColor, box: boxOf(dr) }));
  const plusHasPic = drafts.some((dr) => dr.kind === "pic");
  const gdi = gdiTextDrafts(emf, effOf);
  drafts.push(
    ...gdi.filter((g) => {
      if (g.kind === "pic") return !plusHasPic;
      if (g.kind === "path") {
        const box = boxOf(g);
        return !plusPaints.some(
          (p) =>
            p.fill === g.fill &&
            p.stroke === g.strokeColor &&
            Math.abs(p.box.x0 - box.x0) <= 2 &&
            Math.abs(p.box.y0 - box.y0) <= 2 &&
            Math.abs(p.box.x1 - box.x1) <= 2 &&
            Math.abs(p.box.y1 - box.y1) <= 2,
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
  const end = Math.min(view.byteLength, d + rs);
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

function pushPath(
  drafts: Draft[],
  rawCmds: PathCmds,
  xf: Xform,
  paint: { fill?: string; strokeColor?: string; strokeWidth?: number; dash?: string },
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
      ? {
          strokeColor: paint.strokeColor,
          strokeWidth: paint.strokeWidth,
          ...(paint.dash ? { dash: paint.dash } : {}),
        }
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
const EMR_CREATE_PEN = 38;
const EMR_CREATE_BRUSH_INDIRECT = 39;
const EMR_DELETE_OBJECT = 40;
const EMR_SETTEXTCOLOR = 24;
const EMR_SETTEXTALIGN = 22;
const EMR_BITBLT = 76;
const EMR_STRETCHDIBITS = 81;
const EMR_EXT_CREATE_FONT = 82;
const EMR_EXT_TEXT_OUT_A = 83;
const EMR_EXT_TEXT_OUT_W = 84;
const EMR_POLYLINE16 = 87;
const EMR_EXT_CREATE_PEN = 95;
// Path construction and consumption: a figure opened with BeginPath grows
// through MoveTo/LineTo/PolylineTo16/PolyBezierTo16 (+ CloseFigure), is
// frozen by EndPath, and paints via FillPath/StrokePath/StrokeAndFillPath.
const EMR_MOVE_TO_EX = 27;
const EMR_BEGIN_PATH = 58;
const EMR_END_PATH = 60;
const EMR_CLOSE_FIGURE = 61;
const EMR_FILL_PATH = 62;
const EMR_STROKE_AND_FILL_PATH = 63;
const EMR_STROKE_PATH = 64;
const EMR_ABORT_PATH = 68;
const EMR_LINE_TO = 54;
const EMR_POLYGON16 = 86;
const EMR_POLYBEZIERTO16 = 88;
const EMR_POLYLINETO16 = 89;

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
/** Replay the carrier's GDI records into the same display space the EMF+ drafts
 *  occupy: its text chain (ExtTextOutW runs, the dual-layer's real text) plus
 *  its polyline strokes (the row underlines of list pages — no EMF+ equivalent
 *  exists for them). Each run's placement comes from the carrier's live world
 *  transform, which folds every domain onto the EMF's device rectangle
 *  (rclBounds) shared with the EMF+ output. Fonts and pens are object-table
 *  based (CreateFontIndirectW/CreatePen define a slot, SelectObject activates
 *  it); text color rides on SetTextColor's COLORREF, stroke color on the pen.
 *  `effOf` folds a parent nesting basis into that transform (nested carriers). */
function gdiTextDrafts(emf: Uint8Array, effOf?: (xf: Xform) => Xform): Draft[] {
  const view = new DataView(emf.buffer, emf.byteOffset);
  const drafts: Draft[] = [];
  const fonts = new Map<number, GdiFont>();
  const pens = new Map<number, { color: string; width: number }>();
  const brushes = new Map<number, string>();
  let selected = -1;
  let selectedPen = -1;
  let selectedBrush = -1;
  let xf: Xform = { ...IDENTITY };
  let textColor: string | undefined;
  // SetTextAlign flags; the GDI device default is TA_TOP (0) — the reference y
  // is the cell top, not the baseline (misreading it sinks every text box).
  let textAlign = 0;
  // The path under construction (figures as raw command tuples), started by
  // BeginPath and frozen by EndPath into `openPath` for a later fill/stroke.
  let figure: PathCmds | null = null;
  let openPath: PathCmds | null = null;
  // SaveDC snapshots the full drawing state we track.
  const states: Array<{
    xf: Xform;
    selected: number;
    selectedPen: number;
    selectedBrush: number;
    textColor?: string;
    textAlign: number;
    figure: PathCmds | null;
    openPath: PathCmds | null;
  }> = [];
  /** Transform and emit the frozen path. Fill comes from the selected brush,
   *  stroke from the selected pen (the wavy panel shapes of corpus files ride
   *  exactly this GDI chain — their EMF+ counterparts draw only fragments). */
  const paintOpenPath = (mode: "fill" | "stroke" | "both"): void => {
    if (!openPath || openPath.length < 2) {
      openPath = null;
      return;
    }
    const eff = effOf ? effOf(xf) : xf;
    const fill = mode !== "stroke" ? brushes.get(selectedBrush) : undefined;
    const pen = mode !== "fill" ? pens.get(selectedPen) : undefined;
    if (fill || pen) {
      // The pen width rides the world transform like text does.
      const scale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
      pushPath(drafts, openPath, eff, {
        ...(fill ? { fill } : {}),
        ...(pen ? { strokeColor: pen.color, strokeWidth: Math.max(pen.width * scale, 1) } : {}),
      });
    }
    openPath = null;
  };
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
          // A leading '@' names GDI's vertical variant of the face — a GDI-only
          // convention browser font matching cannot resolve (it falls back to a
          // thinner face), and the vertical geometry comes from the world
          // transform anyway, so strip it for the rendered family.
          face: face.replace(/^@/, ""),
        });
        break;
      }
      case EMR_SELECT_OBJECT: {
        // Stock objects carry a flag bit; real slots reference our tables.
        const raw = view.getUint32(eo + 8, true);
        if (!(raw & 0x80000000)) {
          if (fonts.has(raw)) selected = raw;
          if (pens.has(raw)) selectedPen = raw;
          if (brushes.has(raw)) selectedBrush = raw;
        }
        break;
      }
      case EMR_DELETE_OBJECT: {
        const slot = view.getUint32(eo + 8, true);
        fonts.delete(slot);
        pens.delete(slot);
        brushes.delete(slot);
        break;
      }
      case EMR_CREATE_BRUSH_INDIRECT: {
        // [ihBrush u32][LOGBRUSH style u32][COLORREF u32] — only the solid
        // style (0) names a paintable flat color; hatched/pattern brushes
        // carry payload elsewhere and are skipped.
        const slot = view.getUint32(eo + 8, true);
        if (view.getUint32(eo + 12, true) === 0) {
          brushes.set(slot, rgbHex(view.getUint32(eo + 16, true)));
        }
        break;
      }
      case EMR_CREATE_PEN: {
        // CREATEPEN: [ihPen][LOGPEN style u32][width POINT ×2][COLORREF].
        const color = rgbHex(view.getUint32(eo + 24, true));
        pens.set(view.getUint32(eo + 8, true), { color, width: view.getUint32(eo + 16, true) });
        break;
      }
      case EMR_EXT_CREATE_PEN: {
        // EXTCREATEPEN: [ihPen][offBmi][cbBmi][offBits][cbBits] then the
        // EXTLOGPEN: [style][width][brushStyle][COLORREF][hatch][entries…].
        pens.set(view.getUint32(eo + 8, true), {
          color: rgbHex(view.getUint32(eo + 40, true)),
          width: view.getUint32(eo + 32, true),
        });
        break;
      }
      case EMR_POLYLINE16: {
        // [bounds 4×i32][count u32][points int16 pairs] — the carrier draws its
        // list-row underlines as these two-point strokes.
        const pen = pens.get(selectedPen);
        if (!pen) break;
        const count = view.getUint32(eo + 24, true);
        if (!count || count > 10_000 || eo + 28 + count * 4 > eo + size) break;
        const eff = effOf ? effOf(xf) : xf;
        const cmds: PathCmds = [];
        for (let i = 0; i < count; i++) {
          const x = view.getInt16(eo + 28 + i * 4, true);
          const y = view.getInt16(eo + 28 + i * 4 + 2, true);
          cmds.push([i === 0 ? "M" : "L", [x, y]]);
        }
        // The pen width rides the world transform like text does.
        const scale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
        pushPath(drafts, cmds, eff, {
          strokeColor: pen.color,
          strokeWidth: Math.max(pen.width * scale, 1),
        });
        break;
      }
      case EMR_STRETCHDIBITS:
      case EMR_BITBLT: {
        // The carrier's own photo blits join the replay as pictures: dual-mode
        // files keep their photos here when the EMF+ layer references image
        // slots that are never defined (its draw is then a no-op). STRETCHDIBITS
        // reads dest @24..28, source @32..44, BITMAPINFO @48..64, rop @68 and
        // dest size @72..76; BITBLT packs dest @24..36, rop @40, source @44..48
        // and its (XformSrc-padded) BITMAPINFO @84..100.
        const isDib = type === EMR_STRETCHDIBITS;
        const rop = view.getUint32(eo + (isDib ? 68 : 40), true);
        // The raster-op opcode is the low byte: 0x20 SRCCOPY, 0x86 SRCPAINT,
        // 0xC6 SRCAND. Masked icon pairs blit a 1bpp white-shape layer with
        // SRCPAINT then a color layer with SRCAND — exactly a screen pass
        // (black regions keep the destination) followed by a multiply pass
        // (white regions keep it), landing the color content inside the shape.
        const low = rop & 0xff;
        if (low !== 0x20 && low !== 0x86 && low !== 0xc6) break;
        const offBmi = view.getUint32(eo + (isDib ? 48 : 84), true);
        const cbBmi = view.getUint32(eo + (isDib ? 52 : 88), true);
        const offBits = view.getUint32(eo + (isDib ? 56 : 92), true);
        const cbBits = view.getUint32(eo + (isDib ? 60 : 96), true);
        if (!cbBmi || cbBmi > 1_000_000 || !cbBits || cbBits > 12_000_000) break;
        if (offBmi < 8 || offBits < offBmi + cbBmi || eo + offBits + cbBits > eo + size) break;
        const bmiW = view.getInt32(eo + offBmi + 4, true);
        const bmiH = Math.abs(view.getInt32(eo + offBmi + 8, true));
        const dx = view.getInt32(eo + 24, true);
        const dy = view.getInt32(eo + 28, true);
        const dw = view.getUint32(eo + (isDib ? 72 : 32), true);
        const dh = view.getUint32(eo + (isDib ? 76 : 36), true);
        const sx = view.getInt32(eo + (isDib ? 32 : 44), true);
        const sy = view.getInt32(eo + (isDib ? 36 : 48), true);
        let crop: PicDraft["crop"];
        if (isDib) {
          const cxS = view.getUint32(eo + 40, true);
          const cyS = view.getUint32(eo + 44, true);
          if (
            bmiW > 0 &&
            bmiH > 0 &&
            cxS > 0 &&
            cyS > 0 &&
            (sx > 0 || sy > 0 || sx + cxS < bmiW || sy + cyS < bmiH)
          ) {
            crop = {
              l: Math.max(0, sx / bmiW),
              t: Math.max(0, sy / bmiH),
              r: Math.max(0, 1 - (sx + cxS) / bmiW),
              b: Math.max(0, 1 - (sy + cyS) / bmiH),
            };
            if (!(crop.l > 0.001 || crop.t > 0.001 || crop.r > 0.001 || crop.b > 0.001))
              crop = undefined;
          }
        }
        const eff = effOf ? effOf(xf) : xf;
        const [px, py] = xformPoint(eff, dx, dy);
        const [px2] = xformPoint(eff, dx + dw, dy);
        const [, py2] = xformPoint(eff, dx, dy + dh);
        if (px2 - px <= 0 || py2 - py <= 0) break;
        drafts.push({
          kind: "pic",
          x: px,
          y: py,
          w: px2 - px,
          h: py2 - py,
          src: bmpDataUrl(emf, eo + offBmi, cbBmi + cbBits),
          ...(low !== 0x20
            ? { blend: low === 0x86 ? ("screen" as const) : ("multiply" as const) }
            : {}),
          ...(crop ? { crop } : {}),
        });
        break;
      }
      case EMR_SETTEXTCOLOR: {
        // COLORREF (0x00bbggrr); high-flag words reference the system palette
        // — skip those and keep the last true RGB.
        const ref = view.getUint32(eo + 8, true);
        if (ref <= 0xffffff) textColor = rgbHex(ref);
        break;
      }
      case EMR_SETTEXTALIGN:
        textAlign = view.getUint32(eo + 8, true);
        break;
      case EMR_SAVEDC:
        states.push({
          xf,
          selected,
          selectedPen,
          selectedBrush,
          textColor,
          textAlign,
          figure,
          openPath,
        });
        break;
      case EMR_RESTOREDC: {
        const st = states.pop();
        if (st) {
          ({ xf, selected, selectedPen, selectedBrush, textColor, textAlign, figure, openPath } =
            st);
        }
        break;
      }
      case EMR_MOVE_TO_EX: {
        // A move starts a figure: corpus exporters skip the BeginPath bracket
        // and grow paths straight from MoveToEx, so a move with no pending
        // figure opens one. With a figure pending the move opens a NEW SUBPATH
        // of the same path — GDI fills all its figures together (the ring
        // icons: two MoveTo'd circles consumed by one FillPath, ALTERNATE
        // winding carving the hole). Replacing there dropped the outer circle
        // and left the inner twin as a solid disc painted over the ring.
        const pt: [number, number] = [view.getInt32(eo + 8, true), view.getInt32(eo + 12, true)];
        figure = figure ? [...figure, ["M", pt]] : [["M", pt]];
        break;
      }
      case EMR_LINE_TO: {
        if (figure) figure.push(["L", [view.getInt32(eo + 8, true), view.getInt32(eo + 12, true)]]);
        break;
      }
      case EMR_POLYGON16: {
        // A standalone closed polygon (same body shape as POLYLINE16),
        // filled with the selected brush and outlined with the pen — the
        // icon artwork (circles, arrows, bar-chart bars) rides these.
        const fill = brushes.get(selectedBrush);
        if (!fill) break;
        const count = view.getUint32(eo + 24, true);
        if (!count || count > 10_000 || eo + 28 + count * 4 > eo + size) break;
        const cmds: PathCmds = [];
        for (let i = 0; i < count; i++) {
          const x = view.getInt16(eo + 28 + i * 4, true);
          const y = view.getInt16(eo + 28 + i * 4 + 2, true);
          cmds.push([i === 0 ? "M" : "L", [x, y]]);
        }
        cmds.push(["Z", []]);
        const pen = pens.get(selectedPen);
        const eff = effOf ? effOf(xf) : xf;
        const scale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
        pushPath(drafts, cmds, eff, {
          fill,
          ...(pen && pen.width > 0
            ? { strokeColor: pen.color, strokeWidth: Math.max(pen.width * scale, 1) }
            : {}),
        });
        break;
      }
      case EMR_POLYLINETO16:
      case EMR_POLYBEZIERTO16: {
        // Both share the POLYLINE16 body: [bounds 4×i32][count u32][points].
        // PolylineTo16 appends line vertices; PolyBezierTo16 appends bezier
        // triplets ([control1, control2, end] per segment, running position
        // opens the segment — the same convention decodePath applies).
        const count = view.getUint32(eo + 24, true);
        if (!count || count > 10_000 || eo + 28 + count * 4 > eo + size || !figure) break;
        if (type === EMR_POLYLINETO16) {
          for (let i = 0; i < count; i++) {
            figure.push([
              "L",
              [view.getInt16(eo + 28 + i * 4, true), view.getInt16(eo + 28 + i * 4 + 2, true)],
            ]);
          }
        } else {
          // PolyBezierTo16 points are [control1, control2, end] triplets;
          // each segment's start anchor is the running position (the previous
          // command's last point), so a straight emit of cubics needs no
          // state. A partial trailing triplet is not a segment.
          for (let i = 0; i + 2 < count; i += 3) {
            const at = eo + 28 + i * 4;
            const p = (o: number): number => view.getInt16(at + o, true);
            figure.push(["C", [p(0), p(2), p(4), p(6), p(8), p(10)]]);
          }
        }
        break;
      }
      case EMR_BEGIN_PATH:
        figure = [];
        openPath = null;
        break;
      case EMR_CLOSE_FIGURE:
        if (figure?.length) figure.push(["Z", []]);
        break;
      case EMR_END_PATH:
        if (figure && figure.length >= 2) openPath = figure;
        figure = null;
        break;
      case EMR_ABORT_PATH:
        figure = null;
        openPath = null;
        break;
      case EMR_FILL_PATH:
        paintOpenPath("fill");
        break;
      case EMR_STROKE_PATH:
        paintOpenPath("stroke");
        break;
      case EMR_STROKE_AND_FILL_PATH:
        paintOpenPath("both");
        break;
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
          // its uniform factor. GDI lfHeight semantics: negative = character
          // em height, positive = cell height (em + internal leading). Glyphs
          // render at the em, so a positive height shrinks by the CJK
          // cell/em ratio — the corpus draws with SimSun/雅黑, whose OS/2
          // winAscent+winDescent ≈ 1.297 em; files without a font record draw
          // stock text at the same cell-scale magnitude.
          const eff = effOf ? effOf(xf) : xf;
          const x = view.getInt32(eo + EXT_REF_X, true);
          const yBaseline = view.getInt32(eo + EXT_REF_Y, true);
          // Column norms survive rotated world transforms: a 90°-rotated
          // (vertical banner) transform zeroes |m11|+|m22|, which used to
          // collapse those runs to height 0 — invisible on the page.
          const glyphScale = Math.max(Math.hypot(eff.m11, eff.m21), Math.hypot(eff.m12, eff.m22));
          const height =
            (font
              ? font.height < 0
                ? -font.height
                : font.height / GDI_CELL_PER_EM
              : 100 / GDI_CELL_PER_EM) * glyphScale;
          // SetTextAlign semantics for the reference y (same calibration as the
          // WMF player): TA_BASELINE hoists by the typical CJK ascent, TA_BOTTOM
          // by the full em, and the device default TA_TOP already names the cell
          // top — hoisting there sank every box by 0.8 em. Advance prefers the
          // trail Dx run (GDI's own per-char advances — tracked labels report
          // wider steps than the glyphs are wide, and the difference is the
          // letter spacing); per-char estimates stand in when it is absent.
          let advance = 0;
          let spacing: number | undefined;
          const dxAdvances = readDxAdvances(
            view,
            strAt + chars * (type === EMR_EXT_TEXT_OUT_W ? 2 : 1),
            chars,
          );
          if (dxAdvances) {
            const natural = Array.from(text).map(
              (ch) => height * (ch.charCodeAt(0) > 0xff ? 1 : 0.55),
            );
            advance = dxAdvances.reduce((sum, a) => sum + a * glyphScale, 0);
            const drift = dxAdvances.reduce(
              (sum, a, i) => sum + a * glyphScale - (natural[i] ?? 0),
              0,
            );
            if (Math.abs(drift) > 0.3 * dxAdvances.length) spacing = drift / dxAdvances.length;
          } else {
            for (const ch of text) advance += height * (ch.charCodeAt(0) > 0xff ? 1 : 0.55);
          }
          const refY = x * eff.m12 + yBaseline * eff.m22 + eff.dy;
          // The text-direction column (m11,m12) names the screen angle the
          // run extends along: 0° horizontal, ±90° vertical (plan-box
          // columns). Unrotated runs stay unmarked.
          const angle = (Math.atan2(eff.m12, eff.m11) * 180) / Math.PI;
          // GDI plays rotated text with upright glyphs — the world transform
          // steers the run's advance but never turns the glyph outlines
          // (PlayEnhMetaFile-verified: the corpus vertical-banner columns
          // render as upright stacked characters, and Word matches). A ±90°
          // run therefore emits one box per character stacked along the
          // advance vector; other angles keep a rotation-carrying box.
          if (Math.abs(Math.abs(angle) - 90) < 0.5) {
            // Advance (logical +x) and cross (logical +y) device columns,
            // normalized to unit logical steps.
            const ua = eff.m11 / glyphScale;
            const ub = eff.m12 / glyphScale;
            const va = eff.m21 / glyphScale;
            const vb = eff.m22 / glyphScale;
            // A vertical (@font) run's reference point names the FIRST CELL's
            // top corner — layout-origin semantics, not GDI's SetTextAlign:
            // hoisting by the TA_BASELINE ascent (0.8 em) lifts every column
            // a full 19px above Word's render. Pixel-verified on all three
            // corpus @楷体 instances (P27/P28/P42, −360 request, TA_BASELINE):
            // with the hoist every column sits exactly 0.8 em high; without
            // it the per-char pitch already matched (27px) and the tops land
            // on the reference. Word's upright replay of vertical runs is the
            // fidelity target; raw GDI would draw them sideways anyway.
            // The cell spans [ref − descent, ref + ascent] across the column
            // (GDI_V_* measured on the realized font), the em box centered in
            // it — pixel-verified against the reference render (ink center
            // lands (ascent − descent)/2 right of the reference).
            const cellL = height * GDI_V_DESCENT_PER_EM;
            const cellR = height * GDI_V_ASCENT_PER_EM;
            const cellInset = (cellL + cellR - height) / 2;
            let prefix = 0;
            let ci = 0;
            for (const ch of text) {
              const step = dxAdvances
                ? dxAdvances[ci] * glyphScale
                : height * (ch.charCodeAt(0) > 0xff ? 1 : 0.55);
              const ax = x * eff.m11 + yBaseline * eff.m21 + eff.dx + ua * prefix;
              const ay = refY + ub * prefix;
              const bx = ax + ua * step;
              const by = ay + ub * step;
              // Corner fold of the advance span against both cross edges.
              const xsn = [
                ax,
                bx,
                ax + va * cellL,
                bx + va * cellL,
                ax - va * cellR,
                bx - va * cellR,
              ];
              const ysn = [
                ay,
                by,
                ay + vb * cellL,
                by + vb * cellL,
                ay - vb * cellR,
                by - vb * cellR,
              ];
              drafts.push({
                kind: "text",
                x: Math.min(...xsn),
                y: Math.min(...ysn),
                w: Math.max(...xsn) - Math.min(...xsn) + 2,
                h: Math.max(...ysn) - Math.min(...ysn),
                text: ch,
                family: font?.face ?? "",
                sizeWorld: height,
                ...(isVerticalPunct(ch) ? { rotation: 90 } : {}),
                ...(!isVerticalPunct(ch) && cellInset > 0.01 ? { cellInsetWorld: cellInset } : {}),
                ...(textColor ? { color: textColor } : {}),
                ...(font && font.weight >= 700 ? { bold: true } : {}),
              });
              prefix += step;
              ci++;
            }
          } else {
            drafts.push({
              kind: "text",
              x: x * eff.m11 + yBaseline * eff.m21 + eff.dx,
              y:
                (textAlign & 0x18) === 0x18
                  ? refY - height * 0.8
                  : (textAlign & 0x18) === 0x08
                    ? refY - height
                    : refY,
              w: advance + 2,
              h: height * 1.35,
              text,
              family: font?.face ?? "",
              sizeWorld: height,
              ...(spacing ? { letterSpacingWorld: spacing } : {}),
              ...(Math.abs(angle) > 0.5 ? { rotation: angle } : {}),
              ...(textColor ? { color: textColor } : {}),
              // GDI face matching splits at FW_BOLD (700): a 681-weight request
              // (corpus norm for 微软雅黑) still resolves to the regular face,
              // and bolding it sinks the label look of the source rendering.
              ...(font && font.weight >= 700 ? { bold: true } : {}),
            });
          }
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

/** Char advances from the trail Dx run of an ExtTextOut record, in record
 *  logical units. Corpus exporters emit them as pseudo-PDY pairs
 *  ([advance, 0] per char), sometimes with one leading straggler word in
 *  front of the first pair; the pair offset that makes every second word
 *  zero wins. No Dx run or no clean pairing → undefined (width falls back
 *  to the per-char estimate). */
function readDxAdvances(view: DataView, at: number, chars: number): number[] | undefined {
  const n = Math.min(chars, 4096);
  if (n < 1) return undefined;
  const word = (i: number): number | undefined =>
    at + i * 2 + 2 <= view.byteLength ? view.getInt16(at + i * 2, true) : undefined;
  for (const off of [0, 1]) {
    const adv: number[] = [];
    let clean = true;
    for (let i = 0; i < n; i++) {
      const a = word(off + i * 2);
      const u = word(off + i * 2 + 1);
      if (a == null || u == null || u !== 0 || a <= 0) {
        clean = false;
        break;
      }
      adv.push(a);
    }
    if (clean) return adv;
  }
  return undefined;
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
 *  the renderer-facing members. Word maps the EMF's declared physical frame
 *  (EMR_HEADER rclFrame at 96 dpi) onto the extent independently per axis —
 *  content past the frame edges stays inside the box instead of stretching
 *  everything above it. When the frame is unavailable or degenerate, the
 *  union bounding box stands in.
 *
 *  An a:srcRect crops the frame first (pixel-verified on the corpus stack
 *  where the visible frame is the top third of the carrier: Word stretches
 *  only that region onto the extent, near-aspect-preserved, and the records
 *  below the crop line never draw — members may now reach past the box, so
 *  the painter clips the replay to the extent like GDI does). */
function finalize(
  drafts: Draft[],
  boxW: number,
  boxH: number,
  frame?: { x: number; y: number; w: number; h: number },
  crop?: SourceCrop,
): MetafileMember[] | undefined {
  let minX: number, minY: number, sX: number, sY: number;
  if (frame && frame.w > 0 && frame.h > 0) {
    minX = frame.x;
    minY = frame.y;
    let fw = frame.w,
      fh = frame.h;
    if (crop) {
      minX += crop.left * fw;
      minY += crop.top * fh;
      fw *= 1 - crop.left - crop.right;
      fh *= 1 - crop.top - crop.bottom;
    }
    if (!(fw > 0 && fh > 0)) return undefined;
    sX = boxW / fw;
    sY = boxH / fh;
  } else {
    let x0 = Infinity,
      y0 = Infinity,
      x1 = -Infinity,
      y1 = -Infinity;
    for (const dr of drafts) {
      if (!participatesInBox(dr)) continue;
      if (dr.kind === "path") {
        for (const [, nums] of dr.cmds) {
          for (let i = 0; i + 1 < nums.length; i += 2) {
            x0 = Math.min(x0, nums[i]);
            x1 = Math.max(x1, nums[i]);
            y0 = Math.min(y0, nums[i + 1]);
            y1 = Math.max(y1, nums[i + 1]);
          }
        }
        continue;
      }
      x0 = Math.min(x0, dr.x);
      y0 = Math.min(y0, dr.y);
      x1 = Math.max(x1, dr.x + dr.w);
      y1 = Math.max(y1, dr.y + dr.h);
    }
    if (crop) {
      const w = x1 - x0,
        h = y1 - y0;
      x0 += crop.left * w;
      y0 += crop.top * h;
      x1 -= crop.right * w;
      y1 -= crop.bottom * h;
    }
    minX = x0;
    minY = y0;
    if (!(x1 - x0 > 0 && y1 - y0 > 0)) return undefined;
    sX = boxW / (x1 - x0);
    sY = boxH / (y1 - y0);
  }
  const X = (x: number) => (x - minX) * sX;
  const Y = (y: number) => (y - minY) * sY;
  const members: MetafileMember[] = [];
  for (const dr of drafts) {
    if (dr.kind === "text") {
      members.push({
        kind: "textBox",
        x: X(dr.x),
        y: Y(dr.y),
        width: dr.w * sX,
        height: dr.h * sY,
        nowrap: true,
        ...(dr.cellInsetWorld ? { insets: { left: dr.cellInsetWorld * sX } } : {}),
        ...(dr.rotation ? { rotation: dr.rotation } : {}),
        runs: [
          {
            text: dr.text,
            family: dr.family,
            sizePx: dr.sizeWorld * sY,
            ...(dr.color ? { color: dr.color } : {}),
            ...(dr.bold ? { bold: true } : {}),
            ...(dr.letterSpacingWorld ? { letterSpacingPx: dr.letterSpacingWorld * sX } : {}),
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
        ...(dr.blend ? { blend: dr.blend } : {}),
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
        ? {
            line: { px: strokeWidth, color: dr.strokeColor, ...(dr.dash ? { dash: dr.dash } : {}) },
          }
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
