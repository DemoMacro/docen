// Layout results — what the engine produces for a block. Coordinates are
// block-relative (a paragraph's line Y starts at its content top; a cell's
// blocks start inside its insets). The flow (page boxing) and the renderer
// (LeaferJS) consume these; `endInlineIndex` marks the source inline a line's
// content ends at (a coarse split-point for page breaking).

import type {
  LayoutBorderEdge,
  LayoutCellInsets,
  LayoutIndent,
  LayoutInline,
  LayoutParagraphBorderEdge,
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

export type LaidOutLineItem = LaidOutTextItem | LaidOutPictureItem;

export interface LaidOutLine {
  yPx: number;
  heightPx: number;
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
}

/** One stacked block with its content-box offset inside the stack (collapsed
 *  before-margins included) — what a renderer needs to place children. */
export interface LaidOutStackItem {
  yPx: number;
  block: LaidOutBlock;
}

export interface LaidOutCell {
  colspan: number;
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
  /** Sum of the spanned columns' widths minus insets/borders — the width the
   *  cell's blocks wrapped at. */
  innerWidthPx: number;
  stack: LaidOutStackItem[];
}

export interface LaidOutRow {
  heightPx: number;
  cells: LaidOutCell[];
}

export interface LaidOutTable {
  kind: "table";
  widthPx: number;
  columnWidthsPx: number[];
  heightPx: number;
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
