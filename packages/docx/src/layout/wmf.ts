// WMF metafile player: replays a placeable-WMF record stream into structured
// scene members (paths, shapes, text boxes, pictures) positioned inside the
// drawing's extent box. This is the vector + text layer the flat DIB
// fallback (wmf-dib.ts) cannot see — Office's ".emf" exports in the wild are
// placeable WMFs whose labels and icon strokes live exactly here.
//
// One pass, a DC state machine, and a GDI object table. Record layouts were
// verified byte-by-byte against the sample corpus: LOGFONT16 is 50 bytes
// (height int16@0, weight@8, italic@10, charset@13, faceName@18 GBK);
// ExtTextOut dx arrays are per-BYTE — a double-byte char advances at its
// lead byte and its trail byte carries 0; point arrays are x,y-ordered
// while scalar params are pushed Y-then-X (the GDI16 calling convention).

import type { LayoutDrawingMember } from "@docen/layout";

import { GDI_CELL_PER_EM, type SourceCrop } from "./emf-plus";
import { bltDibAt, bmpDataUrl } from "./wmf-dib";

const CREATE_PEN = 0x02fa;
const CREATE_BRUSH = 0x02fc;
const CREATE_FONT = 0x02fb;
// Windows.h's META_SELECTOBJECT is 0x012D — 0x0125 is a different record;
// the corpus's create/select/delete triple pins this empirically.
const SELECT_OBJECT = 0x012d;
const DELETE_OBJECT = 0x01f0;
const SAVE_DC = 0x001e;
const RESTORE_DC = 0x0127;
const SET_TEXT_COLOR = 0x0209;
const SET_TEXT_ALIGN = 0x012e;
const SET_POLY_FILL_MODE = 0x0106;
const SET_WINDOW_ORG = 0x020b;
const SET_WINDOW_EXT = 0x020c;
const MOVE_TO = 0x0214;
const LINE_TO = 0x0213;
const POLYGON = 0x0324;
const POLYLINE = 0x0325;
const POLY_POLYGON = 0x0538;
const RECTANGLE = 0x041b;
const ROUND_RECT = 0x061c;
const ELLIPSE = 0x0418;
const EXT_TEXT_OUT = 0x0a32;
const DIB_STRETCH_BLT = 0x0b41;
const STRETCH_DIB = 0x0f43;

const SRCCOPY = 0x00cc0020;
const PS_NULL = 5;
const BS_NULL = 1;
const MEMBERS_CAP = 4000;

/** GDI logical object parked in a slot between SelectObject calls. */
type GdiObject =
  | { kind: "pen"; style: number; width: number; color: number }
  | { kind: "brush"; style: number; color: number }
  | { kind: "font"; height: number; weight: number; italic: boolean; face: string };

interface DcState {
  pen: Extract<GdiObject, { kind: "pen" }> | null;
  brush: Extract<GdiObject, { kind: "brush" }> | null;
  font: Extract<GdiObject, { kind: "font" }> | null;
  textColor: number;
  /** SetTextAlign flags; the GDI device default is TA_TOP (the reference y is
   *  the cell top, not the baseline — misreading it sinks every text box). */
  textAlign: number;
  fillRule: "evenodd" | "nonzero";
  orgX: number;
  orgY: number;
  extX: number;
  extY: number;
  curX: number;
  curY: number;
}

const gbkDecoder = new TextDecoder("gbk");

/** Replay a placeable-WMF byte stream into members laid out inside a
 *  boxW×boxH extent. Returns undefined when the bytes are not a placeable
 *  WMF or produce no drawable member — callers fall back to the flat DIB
 *  picture. An a:srcRect crop selects the window region that stretches onto
 *  the extent (records outside it land past the box for the painter to clip,
 *  matching Word's source-then-stretch semantics). */
export function wmfMembers(
  bytes: Uint8Array,
  boxW: number,
  boxH: number,
  crop?: SourceCrop,
): LayoutDrawingMember[] | undefined {
  if (bytes.length < 22 + 18 + 6) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) !== 0x9ac6cdd7) return undefined;
  // Placeable bbox (Left/Top/Right/Bottom @6/8/10/12) seeds the window; the
  // stream's own SetWindowOrg/Ext records override it afterwards.
  const st: DcState = {
    pen: null,
    brush: null,
    font: null,
    textColor: 0,
    textAlign: 0,
    fillRule: "evenodd",
    orgX: view.getInt16(6, true),
    orgY: view.getInt16(8, true),
    extX: view.getInt16(10, true) - view.getInt16(6, true),
    extY: view.getInt16(12, true) - view.getInt16(8, true),
    curX: 0,
    curY: 0,
  };

  const members: LayoutDrawingMember[] = [];
  const objects: (GdiObject | undefined)[] = [];
  // Create*Indirect parks the object in the lowest free slot — GDI reuses
  // released handles, so a DeleteObject + create pair must not leave the
  // new object dangling at a higher index than the stream's SelectObject.
  const alloc = (obj: GdiObject): void => {
    const slot = objects.indexOf(undefined);
    if (slot >= 0) objects[slot] = obj;
    else objects.push(obj);
  };
  const saved: DcState[] = [];
  // A MoveTo/LineTo run buffers into one open polyline per stroke sequence.
  let lineBuf: Pt[] = [];

  // Logical → box px, against the live window (SetWindowExt may rescale
  // mid-stream). Scales also convert GDI sizes — font height, pen width —
  // into box px. A srcRect narrows the mapped window: the crop fractions
  // ride the live window so a mid-stream SetWindowOrg/Ext keeps working.
  const cropX = (): number => (crop ? 1 - crop.left - crop.right : 1);
  const cropY = (): number => (crop ? 1 - crop.top - crop.bottom : 1);
  const sx = (): number => boxW / (st.extX * cropX());
  const sy = (): number => boxH / (st.extY * cropY());
  const box = (x: number, y: number): Pt => ({
    x: (x - (st.orgX + (crop ? crop.left * st.extX : 0))) * sx(),
    y: (y - (st.orgY + (crop ? crop.top * st.extY : 0))) * sy(),
  });

  const stroke = (): { px: number; color?: string } | undefined => {
    const pen = st.pen;
    if (!pen || pen.style === PS_NULL) return undefined;
    return {
      px: Math.max(1, Math.round(Math.abs(pen.width) * ((sx() + sy()) / 2))),
      color: hexOf(pen.color),
    };
  };
  const fill = (): string | undefined => {
    const brush = st.brush;
    // Hatched/pattern brushes flatten to their base color.
    if (!brush || brush.style === BS_NULL) return undefined;
    return hexOf(brush.color);
  };
  const flushLine = (): void => {
    if (lineBuf.length >= 2) {
      pushPathMember(members, [lineBuf], false, undefined, stroke());
    }
    lineBuf = [];
  };

  let off = 40; // placeable (22) + standard WMF header (18)
  while (off + 6 <= bytes.length) {
    const sizeWords = view.getUint32(off, true);
    if (sizeWords < 3) break;
    const end = off + sizeWords * 2;
    if (end > bytes.length) break;
    const fn = view.getUint16(off + 4, true);
    const p = off + 6;
    switch (fn) {
      case CREATE_PEN: {
        if (end - p >= 10) {
          alloc({
            kind: "pen",
            style: view.getUint16(p, true),
            width: view.getInt32(p + 2, true),
            color: view.getUint32(p + 6, true) & 0xffffff,
          });
        }
        break;
      }
      case CREATE_BRUSH: {
        // LogBrush16: Style(2) + ColorRef(4) + Hatch(2) — 8 bytes.
        if (end - p >= 8) {
          alloc({
            kind: "brush",
            style: view.getUint16(p, true),
            color: view.getUint32(p + 2, true) & 0xffffff,
          });
        }
        break;
      }
      case CREATE_FONT: {
        // LOGFONT16 (50 bytes); faceName is GBK bytes up to the first NUL.
        if (end - p < 50) break;
        let faceEnd = p + 18;
        while (faceEnd < p + 50 && bytes[faceEnd] !== 0) faceEnd++;
        alloc({
          kind: "font",
          height: view.getInt16(p, true),
          weight: view.getUint16(p + 8, true),
          italic: bytes[p + 10] !== 0,
          // A leading '@' names GDI's vertical variant of the face — a GDI-only
          // convention browser font matching cannot resolve, and the vertical
          // geometry comes from the record geometry, so strip it.
          face: gbkDecoder.decode(bytes.subarray(p + 18, faceEnd)).replace(/^@/, ""),
        });
        break;
      }
      case SELECT_OBJECT: {
        if (end - p < 2) break;
        const obj = objects[view.getUint16(p, true)];
        if (obj?.kind === "pen") st.pen = obj;
        else if (obj?.kind === "brush") st.brush = obj;
        else if (obj?.kind === "font") st.font = obj;
        break;
      }
      case DELETE_OBJECT: {
        if (end - p < 2) break;
        // Slots free for reuse: GDI reallocates released handles, so a later
        // Create*Indirect must not land on a stale object.
        objects[view.getUint16(p, true)] = undefined;
        break;
      }
      case SAVE_DC: {
        saved.push({ ...st });
        break;
      }
      case RESTORE_DC: {
        // Negative count pops that many frames and restores the last popped.
        const n = -view.getInt16(p, true);
        for (let i = 0; i < n && saved.length > 0; i++) Object.assign(st, saved.pop());
        break;
      }
      case SET_TEXT_COLOR: {
        if (end - p >= 4) st.textColor = view.getUint32(p, true) & 0xffffff;
        break;
      }
      case SET_TEXT_ALIGN: {
        if (end - p >= 2) st.textAlign = view.getUint16(p, true);
        break;
      }
      case SET_POLY_FILL_MODE: {
        // 1 = ALTERNATE (even-odd), 2 = WINDING (non-zero).
        st.fillRule = view.getUint16(p, true) === 2 ? "nonzero" : "evenodd";
        break;
      }
      case SET_WINDOW_ORG: {
        st.orgY = view.getInt16(p, true);
        st.orgX = view.getInt16(p + 2, true);
        break;
      }
      case SET_WINDOW_EXT: {
        st.extY = view.getInt16(p, true);
        st.extX = view.getInt16(p + 2, true);
        break;
      }
      case MOVE_TO: {
        flushLine();
        st.curY = view.getInt16(p, true);
        st.curX = view.getInt16(p + 2, true);
        lineBuf = [box(st.curX, st.curY)];
        break;
      }
      case LINE_TO: {
        if (lineBuf.length === 0) lineBuf = [box(st.curX, st.curY)];
        st.curY = view.getInt16(p, true);
        st.curX = view.getInt16(p + 2, true);
        lineBuf.push(box(st.curX, st.curY));
        break;
      }
      case POLYGON: {
        flushLine();
        const pts = readPoints(view, p, end);
        if (pts)
          pushPathMember(
            members,
            [pts.map((pt) => box(pt.x, pt.y))],
            true,
            fill(),
            stroke(),
            st.fillRule,
          );
        break;
      }
      case POLYLINE: {
        flushLine();
        const pts = readPoints(view, p, end);
        if (pts)
          pushPathMember(members, [pts.map((pt) => box(pt.x, pt.y))], false, undefined, stroke());
        break;
      }
      case POLY_POLYGON: {
        flushLine();
        const sub = readSubPolygons(view, p, end);
        if (sub)
          pushPathMember(
            members,
            sub.map((s) => s.map((pt) => box(pt.x, pt.y))),
            true,
            fill(),
            stroke(),
            st.fillRule,
          );
        break;
      }
      case RECTANGLE:
      case ROUND_RECT:
      case ELLIPSE: {
        flushLine();
        // Box params are pushed Bottom, Right, Top, Left (RoundRect leads
        // with its ellipse Height, Width); the y-pair can arrive swapped —
        // normalize.
        const lead = fn === ROUND_RECT ? 4 : 0;
        if (end - p < lead + 8) break;
        const b = view.getInt16(p + lead, true);
        const r = view.getInt16(p + lead + 2, true);
        const t = view.getInt16(p + lead + 4, true);
        const l = view.getInt16(p + lead + 6, true);
        const tl = box(l, Math.min(t, b));
        const br = box(Math.max(l, r), Math.max(t, b));
        const width = br.x - tl.x;
        const height = br.y - tl.y;
        if (width <= 0 || height <= 0) break;
        members.push({
          kind: "shape",
          x: tl.x,
          y: tl.y,
          width,
          height,
          preset: fn === ELLIPSE ? "ellipse" : fn === ROUND_RECT ? "roundRect" : "rect",
          fill: fill(),
          line: stroke(),
        });
        break;
      }
      case EXT_TEXT_OUT: {
        flushLine();
        pushTextMember(members, view, bytes, p, end, st, box, sx, sy);
        break;
      }
      case DIB_STRETCH_BLT:
      case STRETCH_DIB: {
        flushLine();
        // 0x0B41: Rop, SrcH, SrcW, YSrc, XSrc, DestH, DestW, YDest, XDest →
        // DIB@20; 0x0F43 inserts ColorUsage after Rop and the DIB sits at 22.
        // Mask layers (SRCPAINT/SRCAND pairs) are an honest miss — SRCCOPY
        // only, the same deal wmf-dib.ts makes.
        if (view.getUint32(p, true) !== SRCCOPY) break;
        const dib = bltDibAt(view, off, end, fn & 0xff);
        if (!dib) break;
        const lead = fn === STRETCH_DIB ? 2 : 0;
        const dh = view.getInt16(p + lead + 12, true);
        const dw = view.getInt16(p + lead + 14, true);
        const dy = view.getInt16(p + lead + 16, true);
        const dx = view.getInt16(p + lead + 18, true);
        if (dw <= 0 || dh <= 0) break;
        const tl = box(dx, dy);
        const br = box(dx + dw, dy + dh);
        members.push({
          kind: "picture",
          x: tl.x,
          y: tl.y,
          width: br.x - tl.x,
          height: br.y - tl.y,
          src: bmpDataUrl(bytes, dib.start, dib.length),
        });
        break;
      }
      default:
        break; // Escape, clip regions, palette/mode sets — no drawable here
    }
    off = end;
    if (members.length > MEMBERS_CAP) return undefined; // degenerate-stream guard
  }
  flushLine();
  return members.length > 0 ? members : undefined;
}

interface Pt {
  x: number;
  y: number;
}

/** COLORREF (0x00BBGGRR) → hex RRGGBB. */
function hexOf(colorRef: number): string {
  return (((colorRef & 0xff) << 16) | (colorRef & 0xff00) | ((colorRef >> 16) & 0xff))
    .toString(16)
    .padStart(6, "0");
}

/** PointS array of Polygon/Polyline: Count(2) then x,y int16 pairs. */
function readPoints(view: DataView, p: number, end: number): Pt[] | undefined {
  const count = view.getUint16(p, true);
  if (count < 2 || p + 2 + count * 4 > end) return undefined;
  const pts: Pt[] = [];
  for (let i = 0; i < count; i++) {
    pts.push({ x: view.getInt16(p + 2 + i * 4, true), y: view.getInt16(p + 4 + i * 4, true) });
  }
  return pts;
}

/** PolyPolygon: polygon count, per-polygon point counts, then the point
 *  arrays back to back. */
function readSubPolygons(view: DataView, p: number, end: number): Pt[][] | undefined {
  const npoly = view.getUint16(p, true);
  if (npoly < 1 || npoly > 64 || p + 2 + npoly * 2 > end) return undefined;
  let q = p + 2 + npoly * 2;
  const sub: Pt[][] = [];
  for (let i = 0; i < npoly; i++) {
    const count = view.getUint16(p + 2 + i * 2, true);
    if (count < 2 || q + count * 4 > end) return undefined;
    const pts: Pt[] = [];
    for (let j = 0; j < count; j++) {
      pts.push({ x: view.getInt16(q + j * 4, true), y: view.getInt16(q + 2 + j * 4, true) });
    }
    q += count * 4;
    sub.push(pts);
  }
  return sub;
}

/** Sub-path point lists → one path member. The member box is the points'
 *  bounding box and the path data is expressed inside it (Leafer paths are
 *  local-space). Returns without emitting when neither fill nor stroke is
 *  active — a fully transparent member would only pollute the scene tree. */
function pushPathMember(
  members: LayoutDrawingMember[],
  subPaths: Pt[][],
  closed: boolean,
  fillColor: string | undefined,
  line: { px: number; color?: string } | undefined,
  fillRule?: "evenodd" | "nonzero",
): void {
  if (!fillColor && !line) return;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const path of subPaths) {
    for (const pt of path) {
      if (pt.x < minX) minX = pt.x;
      if (pt.y < minY) minY = pt.y;
      if (pt.x > maxX) maxX = pt.x;
      if (pt.y > maxY) maxY = pt.y;
    }
  }
  if (!(maxX > minX || maxY > minY)) return;
  let d = "";
  for (const path of subPaths) {
    for (let i = 0; i < path.length; i++) {
      d += `${i === 0 ? "M" : "L"}${round1(path[i].x - minX)},${round1(path[i].y - minY)}`;
    }
    if (closed) d += "Z";
  }
  members.push({
    kind: "path",
    x: minX,
    y: minY,
    width: Math.max(maxX - minX, 0.5),
    height: Math.max(maxY - minY, 0.5),
    d,
    fill: fillColor,
    line,
    fillRule,
  });
}

function round1(n: number): number {
  return Math.round(n * 10) / 10;
}

/** ExtTextOut: Y, X, byte Count, Options (a rect follows when OPAQUE or
 *  CLIPPED), the GBK string, then a per-BYTE dx array — a double-byte
 *  char's advance sits at its lead byte and its trail byte carries 0. The
 *  emitted box is sized to the advance sum so the text never wraps. */
function pushTextMember(
  members: LayoutDrawingMember[],
  view: DataView,
  bytes: Uint8Array,
  p: number,
  end: number,
  st: DcState,
  box: (x: number, y: number) => Pt,
  sx: () => number,
  sy: () => number,
): void {
  if (end - p < 8) return;
  const y = view.getInt16(p, true);
  const x = view.getInt16(p + 2, true);
  const count = view.getUint16(p + 4, true);
  const opts = view.getUint16(p + 6, true);
  let q = p + 8;
  if (opts & 0x06) q += 8; // skip the ETO_OPAQUE/CLIPPED rect
  if (count === 0 || q + count > end) return;
  const font = st.font;
  if (!font) return;
  const text = gbkDecoder.decode(bytes.subarray(q, q + count));
  // GDI lfHeight: negative = em height, positive = cell height (em + internal
  // leading ≈ 1.297 em for the corpus's CJK fonts) — glyphs render at the em.
  const sizePx = (font.height < 0 ? -font.height : font.height / GDI_CELL_PER_EM) * sy();
  if (!text || sizePx <= 0) return;
  // Advance sum over the dx entries — one per byte; a double-byte char's
  // trail entry is a 0 spacer, so a plain sum is exact. Absent or zero dx
  // falls back to the no-dx estimate: CJK cells ~1em, Latin ~0.55em.
  q += count;
  let advancePx = 0;
  for (let i = 0; i < count && q + 2 <= end; i++) {
    advancePx += view.getInt16(q + i * 2, true) * sx();
  }
  if (advancePx <= 0) {
    for (const ch of text) advancePx += sizePx * (ch.charCodeAt(0) > 0xff ? 1 : 0.55);
  }
  // SetTextAlign semantics for the record's reference y: TA_BASELINE hoists by
  // the typical CJK ascent (approximate; visual calibration refines), TA_BOTTOM
  // by the full em, and the device default TA_TOP already names the cell top —
  // hoisting there sank every box by 0.8 em.
  const vertical = st.textAlign & 0x18;
  const origin = box(x, y);
  members.push({
    kind: "textBox",
    x: origin.x,
    y:
      vertical === 0x18
        ? origin.y - sizePx * 0.8
        : vertical === 0x08
          ? origin.y - sizePx
          : origin.y,
    width: Math.ceil(advancePx) + 2,
    height: Math.ceil(sizePx * 1.35),
    nowrap: true,
    blocks: [
      {
        kind: "paragraph",
        inline: [
          {
            kind: "text",
            text,
            style: {
              family: font.face,
              sizePx,
              color: hexOf(st.textColor),
              bold: font.weight >= 550 || undefined,
              italic: font.italic || undefined,
            },
          },
        ],
      },
    ],
  });
}
