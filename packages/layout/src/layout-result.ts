// Layout results — what the engine produces for a block. Coordinates are
// block-relative (a paragraph's line Y starts at its content top; a cell's
// blocks start inside its insets). The flow (page boxing) and the renderer
// (LeaferJS) consume these; `endInlineIndex` marks the source inline a line's
// content ends at (a coarse split-point for page breaking).

import type {
  LayoutBorderEdge,
  LayoutCellInsets,
  LayoutDrawing,
  LayoutIndent,
  LayoutInline,
  LayoutParagraphBorderEdge,
  LayoutTableBorders,
} from "./layout-doc";

export interface LaidOutTextItem {
  kind: "text";
  /** Index into the paragraph's `inline` array. */
  inlineIndex: number;
  /** The run's whitespace-collapsed slice for this line — the string the
   *  renderer paints (leading spaces of soft wraps collapsed, runs joined),
   *  not a verbatim slice of the source run's text. */
  text: string;
  xPx: number;
  widthPx: number;
}

export interface LaidOutPictureItem {
  kind: "picture";
  inlineIndex: number;
  xPx: number;
  widthPx: number;
  heightPx: number;
}

export interface LaidOutTabItem {
  kind: "tab";
  inlineIndex: number;
  /** The tab's advance interval: from the preceding content's end to the
   *  following content's start (a right stop lands the next run's right edge
   *  at the stop). The painter draws the leader fill across it. */
  xPx: number;
  widthPx: number;
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
}

export type LaidOutLineItem = LaidOutTextItem | LaidOutPictureItem | LaidOutTabItem;

export interface LaidOutLine {
  yPx: number;
  heightPx: number;
  /** The line's max text natural height (0 when no text) — the half-leading
   *  reference when a grid span centers its content. */
  naturalPx: number;
  /** The line's largest run font size (undefined on textless lines) — the
   *  EM-box reference Word centers in a grid span (corpus-verified: the honor
   *  table's rows center the 12pt em box, ~10px above in a 34.7px cell line;
   *  the browser font box the natural height measures runs ~0.3em deeper). */
  textEmPx?: number;
  /** The height is docGrid-derived in the body flow: Word centers the natural
   *  box in the span; non-grid slack sinks below the text instead. */
  grid?: boolean;
  /** A picture floored this line's height — a grid line ceils to whole rows
   *  and centers the picture box itself (its height is the natural box). */
  pictureFloored?: boolean;
  /** This line's own first-line indent (w:ind/@w:firstLine, negative for a
   *  hanging indent) — set on the paragraph's FIRST line only, so a split
   * tail's leading line (mid-paragraph, mid-page) carries none. The painter
   * offsets by it instead of guessing from the line index. */
  firstLineIndentPx?: number;
  /** The source inline the line's content ends at — a coarse split-point
   *  marker for page breaking. */
  endInlineIndex: number;
  items: LaidOutLineItem[];
  /** The line's content width (the justification stretch target) — items are
   *  re-spaced so the last one ends here. Undefined on unjustified lines. */
  maxWidthPx?: number;
  /** Per-gap stretch the layout applied (undefined on unjustified lines).
   *  A justified line is any line where this is set — the painter stretches
   *  each item to the next item's x (the last one to `maxWidthPx`). */
  justifyGapPx?: number;
  /** The advance of the closing punctuation hanging past this line's right
   *  edge (w:overflowPunct) — the painter's stretch target for the last item
   *  extends by it, so the full glyphs fill the width and the closer hangs
   *  into the margin at its natural advance. */
  hangPx?: number;
  /** How far the line's content start sits right of the paragraph's text box
   *  edge — set when a wrapSide right/largest float takes the left side and
   *  the text packs past its right edge (the painter shifts the line). */
  xOffsetPx?: number;
}

export interface LaidOutParagraph {
  kind: "paragraph";
  heightPx: number;
  /** Spacing margins for the stacking caller: `beforePx` above, `afterPx`
   *  below (collapse between siblings is the stacker's job). */
  beforePx: number;
  afterPx: number;
  lines: LaidOutLine[];
  /** The input inline atoms, mirrored so the renderer can read the text and
   *  style behind each line item (`inline[item.inlineIndex]`). */
  inline: readonly LayoutInline[];
  /** Pagination controls mirrored from the input — the flow strategy's
   *  split/move decisions read them off the laid tree. */
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  /** Borders mirrored from the input (w:pBdr) — the painter draws them. */
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutParagraphBorderEdge>>;
  /** Indents mirrored from the input — the painter offsets each line's origin
   *  (left on every line, firstLine additionally on line 0; hanging < 0). */
  indent?: LayoutIndent;
  /** Floating drawings anchored to this paragraph, mirrored for the painter
   *  (the flow gives them no height). */
  drawings?: LayoutDrawing[];
}

/** One stacked block with its content-box offset inside the stack (collapsed
 *  before-margins included) — what a renderer needs to place children. */
export interface LaidOutStackItem {
  yPx: number;
  block: LaidOutBlock;
}

export interface LaidOutCell {
  colspan: number;
  /** Grid rows the cell spans (w:vMerge gridSpan resolved by the adapter). */
  rowspan: number;
  /** Effective insets used (cell's own ?? table default, per side). */
  insets: LayoutCellInsets;
  /** Declared border edges, mirrored from the input for the renderer (the
   *  engine only measures them; adjacent-cell collapse is a render concern). */
  borders?: {
    top?: LayoutBorderEdge;
    right?: LayoutBorderEdge;
    bottom?: LayoutBorderEdge;
    left?: LayoutBorderEdge;
  };
  /** Cell shading (hex RRGGBB), mirrored for the renderer. */
  fill?: string;
  /** Sum of the spanned columns' widths minus insets/borders — the width the
   *  cell's blocks wrapped at. */
  innerWidthPx: number;
  /** w:vAlign offset inside the row (the slack above the content when the row
   *  is taller — center/bottom placement). */
  contentOffsetYPx?: number;
  stack: LaidOutStackItem[];
}

export interface LaidOutRow {
  heightPx: number;
  cells: LaidOutCell[];
  /** Leading w:tblHeader row (a contiguous prefix from the first row) — the
   *  flow re-inserts stripped copies when the table splits across pages. */
  tableHeader?: boolean;
  /** w:cantSplit — the row moves whole to the next page instead of splitting
   *  mid-content (unless it is taller than a whole page, where Word force-
   *  splits rather than clip). */
  cantSplit?: boolean;
  /** w:trHeight exact — the row's height is fixed, so mid-content splitting
   *  is meaningless (overflow clips); the flow always moves it whole. */
  exactHeight?: boolean;
}

export interface LaidOutTable {
  kind: "table";
  widthPx: number;
  columnWidthsPx: number[];
  /** w:jc offset from the flow column's left edge (negative when a wider
   *  table centers into the margins). */
  offsetXPx?: number;
  heightPx: number;
  /** Table-level border defaults, mirrored for the renderer (a cell's missing
   *  edge falls back to these per side). */
  borders?: LayoutTableBorders;
  rows: LaidOutRow[];
}

export interface LaidOutGroup {
  kind: "group";
  heightPx: number;
  children: LaidOutStackItem[];
}

/** Mirrors LayoutPlaceholder — geometry only, the renderer draws the label. */
export interface LaidOutPlaceholder {
  kind: "placeholder";
  heightPx: number;
  label?: string;
}

export interface LaidOutPageBreak {
  kind: "pageBreak";
  heightPx: 0;
}

export type LaidOutBlock =
  | LaidOutParagraph
  | LaidOutTable
  | LaidOutGroup
  | LaidOutPlaceholder
  | LaidOutPageBreak;
