// LayoutDoc — the engine's own input projection. Adapters (docx from
// Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) build
// these plain blocks; the engine never sees a Tiptap node or an OOXML part.
//
// Projection contract:
// - All geometry is px. Unit conversions happen once, in the adapter.
// - Style cascades are resolved. `spacing`/`indent`/text styles arrive as the
//   effective values (direct attrs → style chain → document defaults merged);
//   the engine has no styles table and no `styleId` — re-resolving a cascade
//   here would duplicate each format's semantics it must not know about.
// - Only layout semantics keep their OOXML shape (a line rule, a grid pitch,
//   a snap flag) — those are the rules this engine exists to implement.

import type { FontSlots } from "./font";

// ── inline content ──

/** Resolved run styling. `family` may be script-slotted; the engine itemizes
 *  a mixed Latin/CJK run by code point and picks the slot per segment. */
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
  | { kind: "text"; text: string; style: LayoutTextStyle; field?: "page" | "numPages" }
  | { kind: "break" }
  | { kind: "tab"; toPx?: number }
  | {
      kind: "picture";
      widthPx: number;
      heightPx: number;
      src?: string /** Vector replay members (a WMF/EMF metafile source): when present, the
       * renderer paints these instead of loading `src` — same renderer-only
       * passthrough contract; the engine never reads beyond widthPx/heightPx. */;
      members?: LayoutDrawingMember[];
    };

// ── floating drawings (anchored shape groups) ──

/** One absolutely-positioned member of a floating drawing. Coordinates and
 *  sizes are px in the drawing's own box (top-left corner the origin) — the
 *  adapter already resolved the group's child coordinate space (chOff/chExt
 *  scaling). */
export type LayoutDrawingMember =
  | {
      kind: "picture";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Renderer media source (data URL); absent → an empty frame. */
      src?: string;
      /** Mirrored content flips (a:xfrm @flipH/@flipV). */
      flipH?: boolean;
      flipV?: boolean;
      /** Source rectangle crop (a:srcRect) as fractions of the image edge
       *  (0-1, each side inward); the painted region is the remainder. */
      crop?: { left: number; top: number; right: number; bottom: number };
    }
  | {
      kind: "shape";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Preset geometry (a:prstGeom @prst). The renderer maps the presets it
       *  knows and skips the rest; custom geometry stays unprojected. */
      preset?: string;
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Outline stroke: width in px + hex color (absent color → ink). */
      line?: { px: number; color?: string };
    }
  | {
      kind: "path";
      x: number;
      y: number;
      width: number;
      height: number;
      /** SVG path data in box coordinates (0,0 … width,height) — the adapter
       *  scaled the geometry's own space (custGeom path w/h) into the box. */
      d: string;
      /** Fill rule for self-intersecting outlines (wmf SetPolyFillMode). */
      fillRule?: "evenodd" | "nonzero";
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Outline stroke (a:ln): width px + hex color + cap/join/dash. */
      line?: {
        px: number;
        color?: string;
        cap?: "round" | "square" | "flat";
        join?: "round" | "bevel" | "miter";
        dash?: string;
      };
    }
  | {
      kind: "textBox";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Text insets px (wps:bodyPr lIns/tIns/rIns/bIns, DrawingML defaults
       *  applied by the adapter). */
      insets?: { left?: number; top?: number; right?: number; bottom?: number };
      /** Vertical anchoring (bodyPr @anchor). */
      anchor?: "top" | "center" | "bottom";
      /** The box's content as projected blocks — the renderer stacks them
       *  inside the box (the flow never sees them; the drawing wraps none). */
      blocks: LayoutBlock[];
    };

/** A floating shape group (wpg:wgp under a wp:anchor): a drawing box anchored
 *  to its paragraph, members absolutely positioned inside. The flow gives the
 *  anchor paragraph no extra height — the anchor's `relative` axes resolve to
 *  page geometry at paint time (the painter owns the page box).
 *
 *  Word's eight relativeFrom values collapse onto these four semantic axes
 *  per side; the axes the adapter cannot resolve yet (character/line anchor
 *  points, inside/outside under mirrored margins) land on their nearest
 *  axis, a registered gap. */
export interface LayoutDrawingAnchor {
  horizontal: {
    /** column = the content box; leftMargin/rightMargin = its edges; page =
     *  the page box. margin/insideMargin → column, character → column,
     *  outsideMargin → rightMargin (unmirrored). */
    relative: "column" | "leftMargin" | "rightMargin" | "page";
    /** px from the reference box's left edge (posOffset), or a fraction of
     *  the reference box's width (percentOffset / 1000). Exclusive with align. */
    offsetPx?: number;
    percent?: number;
    /** Alignment inside the reference box (inside→left, outside→right). */
    align?: "left" | "center" | "right";
  };
  vertical: {
    /** paragraph = the anchor paragraph's top; topMargin/bottomMargin = the
     *  content box's edges; page = the page box. margin/insideMargin →
     *  topMargin, line → paragraph, outsideMargin → bottomMargin. */
    relative: "paragraph" | "topMargin" | "bottomMargin" | "page";
    offsetPx?: number;
    percent?: number;
    align?: "top" | "center" | "bottom";
  };
}

export interface LayoutDrawing {
  anchor: LayoutDrawingAnchor;
  width: number;
  height: number;
  members: LayoutDrawingMember[];
  /** w:wrap — how text flows around the box. Absent (wrapNone) paints over
   *  or under the text without affecting the flow; "square"/"tight" shrink
   *  the lines the box overlaps (tight reduces to square's rectangle — a
   *  contour pass is a registered gap); "topAndBottom" clears the band. */
  wrap?: "square" | "tight" | "topAndBottom";
  /** w:behindDoc — painted under the text layer (text strokes stay visible). */
  behind?: boolean;
}

// ── blocks ──

/** OOXML w:spacing/@w:line + @w:lineRule, unit-resolved by the adapter. */
export type LayoutLineHeight =
  | { rule: "exact"; px: number }
  | { rule: "atLeast"; px: number }
  | { rule: "multiple"; factor: number }; // w:line in 240ths of a single line

/** Paragraph spacing (w:spacing), all px. `before`/`after` are data for the
 *  flow: vertical margin collapse between siblings is applied by the caller
 *  that stacks blocks (the flow for body paragraphs, the cell for table
 *  content — a cell is a BFC that eats the first `before` and last `after`). */
export interface LayoutSpacing {
  lineHeight?: LayoutLineHeight;
  beforePx: number;
  afterPx: number;
}

/** Paragraph indents (w:ind), px. `firstLinePx` shrinks only the first line's
 *  wrapping width (CSS text-indent); negative values are hanging indents (a
 *  numbering bullet's line reaches left of the body text — the first line is
 *  WIDER than the rest). The adapter resolved twips — and firstLineChars
 *  (chars/100 × font size) — into px. */
export interface LayoutIndent {
  leftPx?: number;
  rightPx?: number;
  firstLinePx?: number;
}

/** One w:tab stop, px from the content-box left edge. */
export interface LayoutTabStop {
  positionPx: number;
  type: "left" | "center" | "right";
  /** w:leader — the fill drawn across the tab's advance ("none" dropped). */
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
}

/** One w:pBdr edge: a border line drawn beside the paragraph. */
export interface LayoutParagraphBorderEdge extends LayoutBorderEdge {
  /** w:space — offset from the text to the line, pt→px. */
  spacePx?: number;
}

export interface LayoutParagraph {
  kind: "paragraph";
  inline: LayoutInline[];
  spacing?: LayoutSpacing;
  indent?: LayoutIndent;
  /** Explicit tab stops (w:tabs), px from the content-box left edge; tabs
   *  beyond the last stop fall to the default grid (720 twips). */
  tabStops?: LayoutTabStop[];
  /** Paragraph borders (w:pBdr) — the painter draws them beside the block. */
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutParagraphBorderEdge>>;
  /** ¶-mark strut size in px (w:pPr/w:rPr/w:sz resolved): an ABSOLUTE line
   *  height for the paragraph-mark line — the sole content of an empty
   *  paragraph, and the minimum of a picture row shorter than a text line. */
  markSizePx?: number;
  /** Default run style (the style chain's run, resolved): the strut font when
   *  the paragraph has no text runs. */
  defaultTextStyle?: LayoutTextStyle;
  /** w:snapToGrid: absent/null = engine default (snap when a grid pitch is
   *  active); explicit false drops the grid pitch. */
  snapToGrid?: boolean | null;
  /** OOXML pagination controls, already resolved through the style cascade
   *  (a heading defaults keepNext=true — the adapter decides that). */
  keepLines?: boolean;
  /** Horizontal alignment (w:jc, resolved through the style cascade). "both"
   *  (justify) stretches every WRAPPED line's inter-character gaps to the
   *  full content width — the paragraph's last line and hard-break lines
   *  keep their natural width; "distribute" stretches every line, the last
   *  included. "center"/"right" shift each line's items by its slack
   *  (trailing whitespace hangs and never counts). */
  align?: "left" | "center" | "right" | "both" | "distribute";
  keepNext?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  /** Floating drawings anchored to this paragraph: wrap-none boxes paint at
   *  their offset; a `wrap` on the drawing also registers a float zone the
   *  flow shrinks lines around (or a cleared band for topAndBottom). */
  drawings?: LayoutDrawing[];
}

/** Cell insets in px: per-cell w:tcMar, or the table's w:tblCellMar default
 *  (cells inherit per side — the adapter or the engine resolves cell ?? table). */
export interface LayoutCellInsets {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/** One border edge of a cell: nil/none/absent sides carry no width. The
 *  visual default border (the DOM route's 1px Table-Grid stamp) is a renderer
 *  decision injected by the adapter — the engine measures only declared edges. */
export interface LayoutBorderEdge {
  style?: string; // "nil"/"none" → no width
  px?: number; // resolved from w:sz (eighths of a point)
  /** Hex RRGGBB (OOXML w:color); absent/auto → the renderer's ink default. */
  color?: string;
}

/** CT_TblBorders: the table-level edges a cell's own w:tcBorders side falls
 *  back to — outer edges for the grid's rim, inside edges for shared lines. */
export interface LayoutTableBorders {
  top?: LayoutBorderEdge;
  bottom?: LayoutBorderEdge;
  left?: LayoutBorderEdge;
  right?: LayoutBorderEdge;
  insideHorizontal?: LayoutBorderEdge;
  insideVertical?: LayoutBorderEdge;
}

export interface LayoutTableCell {
  colspan?: number;
  rowspan?: number;
  /** Per-spanned-column widths in px (w:tcW resolved); absent → grid share. */
  widthPx?: number;
  insets?: LayoutCellInsets;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>;
  /** Cell shading (w:shd @w:fill), hex RRGGBB. */
  fill?: string;
  /** w:vAlign — the content's placement when the row is taller than it. */
  verticalAlign?: "top" | "center" | "bottom";
  blocks: LayoutBlock[];
}

export interface LayoutTableRow {
  cells: LayoutTableCell[];
  /** w:trHeight resolved: atLeast floors the row, exact fixes it (content
   *  overflows but the row does not grow). */
  height?: { rule: "atLeast" | "exact"; px: number };
}

export type LayoutTableWidth =
  | { type: "percent"; percent: number } // 0-100
  | { type: "px"; px: number }; // w:tblW dxa resolved

export interface LayoutTable {
  kind: "table";
  /** Absent width = auto: fill the containing flow width. */
  width?: LayoutTableWidth;
  /** w:tblPr/w:jc — the table box's placement inside the flow column. A table
   *  wider than the column centers into the margins (negative offset). */
  align?: "left" | "center" | "right";
  /** tblGrid column widths in px, scaled proportionally to the effective
   *  table width (Word scales the grid to tblW, never to the raw sum). */
  columnWidthsPx?: number[];
  /** Table-level default insets (w:tblCellMar) a cell without its own w:tcMar
   *  inherits, per side. */
  cellInsets?: LayoutCellInsets;
  /** Table-level border defaults (w:tblBorders, style chain resolved):
   *  the renderer falls a cell's missing edge back to these per side. */
  borders?: LayoutTableBorders;
  rows: LayoutTableRow[];
}

/** A plain container (list body, blockquote) — laid out by recursion, no
 *  geometry of its own. */
export interface LayoutGroup {
  kind: "group";
  blocks: LayoutBlock[];
}

/** An opaque block the adapter cannot lay out (raw XML passthrough, TOC,
 *  altChunk…): the flow reserves an estimated box so the content keeps a
 *  visual presence instead of silently vanishing. The renderer draws the
 *  label; the engine only moves the box. */
export interface LayoutPlaceholder {
  kind: "placeholder";
  /** Estimated height in px (the adapter's guess — usually N default lines). */
  heightPx: number;
  /** What the box stands for, shown by the renderer (e.g. "toc", "rawXml"). */
  label?: string;
}

/** A page/column break atom: zero height, closes the current flow box after
 *  the preceding content (the break never opens a page — Word semantics). */
export interface LayoutPageBreak {
  kind: "pageBreak";
}

export type LayoutBlock =
  | LayoutParagraph
  | LayoutTable
  | LayoutGroup
  | LayoutPlaceholder
  | LayoutPageBreak;

// ── flow context ──

/** A floating picture's flow band: text lines overlapping [topPx, bottomPx)
 *  wrap beside it (usable width reduced by widthPx). wrapNone/page/margin
 *  anchors never produce a zone — they sit outside the text flow. */
export interface LayoutFloatZone {
  widthPx: number;
  topPx: number;
  bottomPx: number;
}

/** Block-level layout context threaded by the caller that stacks blocks. */
export interface LayoutBlockContext {
  /** Section document-grid pitch in px (w:docGrid linePitch); 0/absent = no
   *  grid. CJK lines ceil to a whole pitch multiple; Latin lines floor at
   *  max(natural, pitch); table cells floor at max(natural, pitch). */
  linePitchPx?: number;
  /** True only for the main body flow: its grid-height lines center the
   *  natural text box in the span (the docGrid lattice). Header/footer and
   *  text-box stacks scale heights by the pitch but place text at the line
   *  top — Word keeps the lattice to the body (verified vs Word). */
  onGrid?: boolean;
  /** True inside a table cell: the grid snap floors instead of ceils (the
   *  row's trHeight governs, not the line box). Cells also clear float
   *  zones — a cell's width is its column, not the page flow. */
  inTable?: boolean;
  floatZones?: readonly LayoutFloatZone[];
  /** This block's top Y within the flow — pairs with floatZones to derive
   *  each line's band. */
  startY?: number;
}
