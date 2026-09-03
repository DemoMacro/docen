import type { LayoutDrawing, LayoutFloatZone } from "./drawing";
import type { LayoutInline, LayoutTextStyle } from "./inline";
import type { LayoutTable } from "./table";

/** OOXML w:spacing/@w:line + @w:lineRule, unit-resolved by the adapter. */
export type LayoutLineHeight =
  | { rule: "exact"; px: number }
  | { rule: "atLeast"; px: number }
  | { rule: "multiple"; factor: number }; // w:line in 240ths of a single line

/** Paragraph spacing (w:spacing), all px. `before`/`after` are data for the
 *  flow: vertical margin collapse between siblings is applied by the caller
 *  that stacks blocks (the body flow and table cells alike — the BFC model:
 *  edge margins count, middles collapse at the max). */
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

/** One border edge of a cell: nil/none/absent sides carry no width. The
 *  visual default border (the DOM route's 1px Table-Grid stamp) is a renderer
 *  decision injected by the adapter — the engine measures only declared edges. */
export interface LayoutBorderEdge {
  style?: string; // "nil"/"none" → no width
  px?: number; // resolved from w:sz (eighths of a point)
  /** Hex RRGGBB (OOXML w:color); absent/auto → the renderer's ink default. */
  color?: string;
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
  /** The document's default tab-grid pitch in px (w:defaultTabStop) — the
   *  fallback grid tabs hop to past the last explicit stop. Absent = 720
   *  twips (Word's default). */
  defaultTabStopPx?: number;
  /** Paragraph borders (w:pBdr) — the painter draws them beside the block. */
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutParagraphBorderEdge>>;
  /** Paragraph shading (w:shd @w:fill), hex RRGGBB — the painter fills the
   *  block box with it beneath everything else the paragraph paints. */
  shadingFill?: string;
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
  /** w:suppressLineNumbers — the paragraph's lines render but never count
   *  toward the section's line numbering (w:lnNumType). */
  suppressLineNumbers?: boolean;
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
  /** The paragraph closes its section (carries the sectPr): the painter swaps
   *  its ¶ for Word's "─────分节符(下一页)─────" mark row. The last section's
   *  sectPr rides the body end and never sets this. */
  sectionEnd?: boolean;
  /** Floating drawings anchored to this paragraph: wrap-none boxes paint at
   *  their offset; a `wrap` on the drawing also shrinks the anchor
   *  paragraph's own lines around the box and registers a float zone the
   *  flow applies to later paragraphs (or a cleared band for topAndBottom). */
  drawings?: LayoutDrawing[];
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

/** A page break atom: zero height, closes the current flow box after
 *  the preceding content (the break never opens a page — Word semantics). */
export interface LayoutPageBreak {
  kind: "pageBreak";
}

/** A column break atom (w:br/@w:type="column"): closes the current column
 *  after the preceding content — the flow continues in the next column, or
 *  on a fresh page past the last one. */
export interface LayoutColumnBreak {
  kind: "columnBreak";
}

export type LayoutBlock =
  | LayoutParagraph
  | LayoutTable
  | LayoutGroup
  | LayoutPlaceholder
  | LayoutPageBreak
  | LayoutColumnBreak;

/** Block-level layout context threaded by the caller that stacks blocks. */
export interface LayoutBlockContext {
  /** Section document-grid pitch in px (w:docGrid linePitch); 0/absent = no
   *  grid. Body CJK lines ceil to a whole pitch multiple; Latin lines floor
   *  at max(natural, pitch); table cells never ceil — the pitch scales a
   *  multiple rule and floors a snapped one (adjustLineHeightInTable). */
  linePitchPx?: number;
  /** The body flow centers a grid-height line's natural text box in the
   *  span (the docGrid lattice — Word-verified). Text-box stacks share the
   *  rule for their grid-snapped lines (half-leading like the body);
   *  header/footer stacks are laid with no grid context at all (natural
   *  line heights) and so never set this. */
  onGrid?: boolean;
  /** True inside a table cell: cell lines join the document grid only when
   *  `adjustLinesInTable` is set (the compat flag), and even then the pitch
   *  is a floor, not a row count (the row's trHeight floors separately).
   *  Cells also clear float zones — a cell's width is its column, not the
   *  page flow. */
  inTable?: boolean;
  /** w:adjustLineHeightInTable (settings.xml compat): cell lines join the
   *  document grid. Absent (its OOXML default) leaves cells grid-free —
   *  CJK-Word documents carry the element, most Western ones don't. */
  adjustLinesInTable?: boolean;
  floatZones?: readonly LayoutFloatZone[];
  /** This block's top Y within the flow — pairs with floatZones to derive
   *  each line's band. */
  startY?: number;
}
