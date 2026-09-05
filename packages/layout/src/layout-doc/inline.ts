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
  /** The w:u pattern token (ST_Underline minus "none") — absent = single. */
  underlineStyle?: string;
  /** w:u color, hex RRGGBB — absent = the text color. */
  underlineColor?: string;
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

/** Phonetic-guide annotation riding a text item (w:ruby — Word's 拼音指南).
 *  The annotation paints above the base glyphs at fontSizePx; the base text
 *  is the item's own text (the editor models the guide as a character mark,
 *  so this is pure paint/layout metadata). */
export interface LayoutRuby {
  /** The annotation text (w:rt, flattened). */
  text: string;
  /** The annotation's distribution over the base (ST_RubyAlign token).
   *  center/left/right shift the annotation within the base's width — the
   *  distribute variants fall to center. */
  alignment?: string;
  /** Annotation font size in px (w:hps resolved; Word's default is half the
   *  base size). */
  fontSizePx: number;
}

/** Two-lines-in-one / combined characters (w:eastAsianLayout @w:combine —
 *  Word's 双行合一 / 合并字符). The run's text packs into two half-size lines
 *  inside a normal line box, optionally wrapped in bracket glyphs. */
export interface LayoutCombine {
  /** The upper line's text. */
  first: string;
  /** The lower line's text. */
  second: string;
  /** The bracket pair drawn around both lines (ST_CombineBrackets minus
   *  "none"); the painter draws the matching glyph characters. */
  bracket?: "round" | "square" | "angle" | "curly";
}

/** One inline item of a paragraph. Text, hard breaks, tabs, and inline
 *  pictures are all line-box content — one box packer handles the four (the
 *  unified text+picture breaker the DOM route never had). A picture's `src`
 *  (data URL or object URL) is renderer-only passthrough — the engine measures
 *  the box and never loads it. A tab advances to the next stop: the explicit
 *  `toPx` (a numbering bullet's hop to the body text), the paragraph's
 *  `tabStops`, or the default grid (720 twips). */

export interface LayoutInlineNoteRef {
  kind: "footnote" | "endnote";
  id: number;
  ordinal: number;
}

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
      /** Synthesized paint content with no document-model character behind it
       *  (a numbering bullet's glyph): the editor's caret/selection mapping
       *  must not count it against the paragraph's text positions. */
      synthetic?: boolean;
      /** Note reference metadata (Word's FootnoteReference / EndnoteReference).
       *  The paginator reads this to anchor footnotes to the page where this run lands. */
      noteRef?: LayoutInlineNoteRef;
      /** Phonetic guide (w:ruby): the annotation paints above the base glyphs
       *  and the line's natural height reserves space for it. */
      ruby?: LayoutRuby;
      /** Two-lines-in-one (w:eastAsianLayout): the atom packs its whole text
       *  into two half-size lines — an unbreakable box of the combined width. */
      combine?: LayoutCombine;
    }
  | { kind: "break" }
  | { kind: "tab"; toPx?: number }
  | {
      /** An OMML formula (m:oMath) the engine does not lay out yet: a fixed
       *  unbreakable placeholder box carrying a short structural label (the
       *  renderer draws it as a dashed slot, Word's empty-argument look). */
      kind: "math";
      label: string;
      widthPx: number;
      heightPx: number;
    }
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
      /** Clockwise spin of the box about its center, degrees (a:xfrm @rot) —
       * the extent stays put, the painted content tilts inside it. */
      rotation?: number;
    };
