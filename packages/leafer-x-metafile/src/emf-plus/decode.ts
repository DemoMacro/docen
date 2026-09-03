import type { PathCmds } from "./draft";
import { GDIPLUS_VERSION, OBJ_BRUSH, OBJ_IMAGE, OBJ_PATH, OBJ_PEN } from "./records";

/** ARGB word (0xAARRGGBB) → hex RRGGBB composited over the white page;
 *  undefined when fully transparent. Metafiles wash whole-page tints at
 *  near-zero alpha (a 10/255 red wash rendered opaque covers the art under
 *  it), so partial alpha pre-blends with white instead of riding along. */
export function argbHex(word: number): string | undefined {
  const a = word >>> 24;
  if (a === 0) return undefined;
  const rgb = word & 0xffffff;
  if (a >= 0xfa) return (rgb | 0x1000000).toString(16).slice(1);
  const mix = (c: number) => Math.round((c * a + 255 * (255 - a)) / 255);
  return [mix(rgb >> 16), mix((rgb >> 8) & 0xff), mix(rgb & 0xff)]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("");
}

// ── object decoding ──
// An object's payload starts at its TotalObjectSize word ([0..3]), then the
// shared version stamp ([4..7]); per-type fields begin at [+8]. Large images
// span multiple consecutive EmfPlusObject records — the caller assembles the
// full byte set first.

export interface BrushInfo {
  solid?: string;
}
export interface PenInfo {
  width: number;
  color?: string;
  /** GDI+ DashStyle preset token (a prstDash value) — the pen's line style. */
  dash?: string;
}

/** GDI+ DashStyle enum values → prstDash tokens (the member protocol's dash
 *  vocabulary); Solid stays absent. DashStyleCustom (5) resolves through the
 *  pen's repeat array instead, so index 5 reads past this table. */
const DASH_STYLE_TOKENS = ["", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"];

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
export interface PathInfo {
  cmds: PathCmds;
}

export function decodeObject(
  payload: Uint8Array,
  type: number,
): BrushInfo | PenInfo | PathInfo | ImageInfo | undefined {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const end = payload.length;
  // Payloads start [TotalObjectSize][Version] with type fields at +8; a
  // re-assembled chunk run keeps the exporter's [chunkDataSize][objectTotal]
  // prefix in front of the stamp — locate the stamp to find the field origin.
  if (end < 8) return undefined;
  let base = 8;
  if (
    view.getUint32(4, true) !== GDIPLUS_VERSION &&
    end >= 16 &&
    view.getUint32(8, true) === GDIPLUS_VERSION
  )
    base = 12;
  if (end < base + 8) return undefined;
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
      if (end < base + 16) return undefined;
      const flags = view.getUint32(base + 4, true);
      const width = view.getFloat32(base + 12, true);
      let cursor = base + 16;
      // A truncated record ends the optional walk — the fields behind the
      // missing bytes are unreadable, the width already decoded stands.
      const past = (by: number): boolean => {
        if (cursor + by > end) return true;
        cursor += by;
        return false;
      };
      // OptionalData fields consumed in ascending PenDataFlags bit order
      // ([MS-EMFPLUS] §2.2.2.28): transform(6 REALs), start/end cap, join,
      // miter limit, dash style DWORD, dashed-line cap, dashed-line offset
      // REAL, then the length-prefixed dash array; caps/compound arrays and
      // non-center/alignment words beyond that are skipped unparsed.
      if (flags & 0x0001 && past(24)) return { width };
      if (flags & 0x0002 && past(4)) return { width };
      if (flags & 0x0004 && past(4)) return { width };
      if (flags & 0x0008 && past(4)) return { width };
      if (flags & 0x0010 && past(4)) return { width };
      let style: number | undefined;
      if (flags & 0x0020) {
        if (past(4)) return { width };
        style = view.getUint32(cursor - 4, true);
      }
      if (flags & 0x0040 && past(4)) return { width };
      if (flags & 0x0080 && past(4)) return { width };
      let dashes: number[] | undefined;
      if (flags & 0x0100) {
        if (past(4)) return { width };
        const n = view.getUint32(cursor - 4, true);
        if (n > 0 && n <= 16 && cursor + n * 4 <= end) {
          dashes = [];
          for (let i = 0; i < n; i++) dashes.push(view.getFloat32(cursor + i * 4, true));
        }
        cursor += Math.min(n, 16) * 4;
      }
      if (flags & 0x0200 && past(4)) return { width };
      if (flags & 0x0400) {
        if (past(4)) return { width };
        const n = view.getUint32(cursor - 4, true);
        cursor += Math.min(n, 64) * 4;
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

export interface ImageInfo {
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
