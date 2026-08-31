import { bmpDataUrl } from "../dib";
import type { Draft, PathCmds, PicDraft } from "./draft";
import { pushPath } from "./draft";
import {
  EXT_CHARS,
  EXT_OFF_STRING,
  EXT_REF_X,
  EXT_REF_Y,
  EMR_ABORT_PATH,
  EMR_BEGIN_PATH,
  EMR_BITBLT,
  EMR_CLOSE_FIGURE,
  EMR_CREATE_BRUSH_INDIRECT,
  EMR_CREATE_PEN,
  EMR_DELETE_OBJECT,
  EMR_END_PATH,
  EMR_EXT_CREATE_FONT,
  EMR_EXT_CREATE_PEN,
  EMR_EXT_TEXT_OUT_A,
  EMR_EXT_TEXT_OUT_W,
  EMR_FILL_PATH,
  EMR_LINE_TO,
  EMR_MODIFY_WORLD_TRANSFORM,
  EMR_MOVE_TO_EX,
  EMR_POLYBEZIERTO16,
  EMR_POLYGON16,
  EMR_POLYLINE16,
  EMR_POLYLINETO16,
  EMR_RESTOREDC,
  EMR_SAVEDC,
  EMR_SELECT_OBJECT,
  EMR_SET_WORLD_TRANSFORM,
  EMR_SETTEXTCOLOR,
  EMR_SETTEXTALIGN,
  EMR_STRETCHDIBITS,
  EMR_STROKE_AND_FILL_PATH,
  EMR_STROKE_PATH,
} from "./records";
import { combine, IDENTITY, readCarrierXform, xformPoint, type Xform } from "./xform";

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

interface GdiFont {
  height: number;
  weight: number;
  face?: string;
}

/** Replay the carrier's GDI records into the same display space the EMF+ drafts
 *  occupy: its text chain (ExtTextOutW runs, the dual-layer's real text) plus
 *  its polyline strokes (the row underlines of list pages — no EMF+ equivalent
 *  exists for them). Each run's placement comes from the carrier's live world
 *  transform, which folds every domain onto the EMF's device rectangle
 *  (rclBounds) shared with the EMF+ output. Fonts and pens are object-table
 *  based (CreateFontIndirectW/CreatePen define a slot, SelectObject activates
 *  it); text color rides on SetTextColor's COLORREF, stroke color on the pen.
 *  `effOf` folds a parent nesting basis into that transform (nested carriers). */
export function gdiTextDrafts(emf: Uint8Array, effOf?: (xf: Xform) => Xform): Draft[] {
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
