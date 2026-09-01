import type { FontSlots } from "../font";
import type { LayoutDrawingMember } from "./drawing";

export interface LayoutTextStyle {
  family: string | FontSlots;
  sizePx: number;
  bold?: boolean;
  italic?: boolean;
  /** Text color, hex RRGGBB (w:color; absent/auto → the renderer's ink). */
  color?: string;
  /** Underlined (w:u, any non-none pattern). */
  underline?: boolean;
  /** Struck through (w:strike / w:dstrike). */
  strikethrough?: boolean;
  /** Extra per-character spacing in px (OOXML w:spacing resolved). */
  letterSpacingPx?: number;
  /** Raised/lowered run (w:vertAlign): glyphs paint at the scaled size on a
   *  shifted baseline (Word's FootnoteReference look). Measuring applies the
   *  same scaling — see vertAlignedSizePx. */
  verticalAlign?: "superscript" | "subscript";
  /** Character highlight (w:highlight), the ST_HighlightColor token — the
   *  painter maps it to Word's highlight palette and fills the run's box
   *  beneath the glyphs. */
  highlight?: string;
  /** Character shading (w:shd) fill, #RRGGBB — the arbitrary-color highlight
   *  Word's shading button paints on selected text. A highlight (same box,
   *  Word's palette) wins when both are present, per OOXML precedence. */
  shadingFill?: string;
}

/** a:srcRect crop as fractions of the image edge (0-1, each side inward);
 *  the painted region is the remainder. */
export interface LayoutPictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** One inline item of a paragraph. Text, hard breaks, tabs, and inline
 *  pictures are all line-box content — one box packer handles the four (the
 *  unified text+picture breaker the DOM route never had). A picture's `src`
 *  (data URL or object URL) is renderer-only passthrough — the engine measures
 *  the box and never loads it. A tab advances to the next stop: the explicit
 *  `toPx` (a numbering bullet's hop to the body text), the paragraph's
 *  `tabStops`, or the default grid (720 twips). */

export type LayoutInline =
  /** A `field` marker makes the text a dynamic page-number atom (w:fldSimple /
   *  complexField PAGE / NUMPAGES): the value only exists after pagination, so
   *  `text` is a single-digit placeholder for measuring and the painter swaps
   *  in the real page number. */
  | {
      kind: "text";
      text: string;
      style: LayoutTextStyle;
      field?: "page" | "numPages";
      /** Ids of the comments whose range covers this atom (w:commentRangeStart
       *  /commentRangeEnd): the painter tints the text's box (sorted, unique).
       *  Pure paint metadata — measuring and wrapping ignore it. */
      commentIds?: number[];
    }
  | { kind: "break" }
  | { kind: "tab"; toPx?: number }
  | {
      kind: "picture";
      widthPx: number;
      heightPx: number;
      src?: string;
      /** a:srcRect crop: the flat `src` paints only the visible remainder
       *  (Leafer paints whole sources, so the renderer sub-regions it). */
      crop?: LayoutPictureCrop;
      /** Vector replay members (a WMF/EMF metafile source): when present, the
       * renderer paints these instead of loading `src` — same renderer-only
       * passthrough contract; the engine never reads beyond widthPx/heightPx. */
      members?: LayoutDrawingMember[];
    };
