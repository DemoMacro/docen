import type { LayoutBlock } from "./block";

// ── Page-level projection types (filled by the format adapters' projectors) ──

/** The flow box a section defines, in px: paper size minus margins
 *  (orientation already resolved by the adapter) and the docGrid pitch. */
export interface ProjectedFlowBox {
  pageWidthPx: number;
  pageHeightPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
  /** Content-box origin within the page (margin left/top) — where the flow's
   *  (0,0) sits on paper; the painter anchors page content here. */
  contentLeftPx: number;
  contentTopPx: number;
  linePitchPx?: number;
}

/** Page furniture (headers/footers) projected for painting: the block lists
 *  per slot (already projected like body blocks — the stage lays them out once
 *  at the content width) plus the placement flags read at paint time.
 *  `headerDistancePx`/`footerDistancePx` are w:pgMar's @w:header/@w:footer
 *  (page edge to the header/footer box; 720 twips = Word's default). */
export interface ProjectedPageFurniture {
  header?: LayoutBlock[];
  firstHeader?: LayoutBlock[];
  evenHeader?: LayoutBlock[];
  footer?: LayoutBlock[];
  firstFooter?: LayoutBlock[];
  evenFooter?: LayoutBlock[];
  /** w:titlePg — page 1 uses the `first` slots instead of `default`. */
  titlePage: boolean;
  /** settings' w:evenAndOddHeaders — even pages use the `even` slots. */
  evenAndOddHeaders: boolean;
  headerDistancePx: number;
  footerDistancePx: number;
}

/** Page background projected for painting (w:background): the solid page
 *  color — a VML pattern fill arrives pre-averaged into it. */
export interface ProjectedPageBackground {
  /** On-page color: the plain w:color, or the pattern tile's threads/gaps
   *  mixed by bit coverage. */
  color?: string;
}

/** One side of the projected page borders (w:pgBorders): the stroke a page
 *  paints on that edge. */
export interface ProjectedPageBorder {
  /** ST_Border token — the painter maps it to a CSS border-style. */
  style: string;
  /** Stroke width in px at 100% zoom (w:sz is 1/8 pt; Word's default 4). */
  widthPx: number;
  /** Hex without '#' (absent = Word's auto, painted black). */
  color?: string;
  /** Distance from the offset reference to this border, in pt. */
  spacePt?: number;
}

/** Page borders projected for painting (w:pgBorders): which pages of the
 *  section paint a border, where the border box measures from, and the
 *  per-side strokes. */
export interface ProjectedPageBorders {
  /** Which of the section's pages paint the border (default allPages). */
  display?: "allPages" | "firstPage" | "notFirstPage";
  /** The border box measures from the page edge (space inset, Word's default
   *  24 pt) or from the text margin (the default). */
  offsetFrom?: "page" | "text";
  /** Paint behind intersecting content (w:zOrder=back; default in front). */
  behind?: boolean;
  top?: ProjectedPageBorder;
  right?: ProjectedPageBorder;
  bottom?: ProjectedPageBorder;
  left?: ProjectedPageBorder;
}

/** Line numbering projected for painting (w:lnNumType): body text lines are
 *  counted across the flow (tables and suppressed paragraphs skip) and a
 *  number paints in the margin beside every countBy-th one. */
export interface ProjectedLineNumbers {
  /** Show a number on every countBy-th counted line (default 1 = all). */
  countBy: number;
  /** The first number after each restart (default 1). */
  start: number;
  /** What resets the counter: every page (the OOXML default), every section,
   *  or nothing (continuous through the document). */
  restart: "newPage" | "newSection" | "continuous";
  /** Gap between the text margin and the number's right edge, in px. Null =
   *  w:distance omitted — the OOXML auto placement (the number centers in the
   *  left margin band, where Word puts it). */
  distancePx: number | null;
}

/** Section columns projected for the flow (w:cols): the page's content box
 *  splits into `count` columns the flow fills left to right before paging.
 *  Absent from a section = one full-width column (the default). */
export interface ProjectedColumns {
  /** Column count (w:cols/@w:num, ≥1). */
  count: number;
  /** Gap between neighboring columns, in px (w:cols/@w:space; Word's default
   *  720 twips when omitted). */
  spacePx: number;
  /** Paint a line between columns (w:cols/@w:sep). */
  separate: boolean;
  /** false = `columnsPx` carries explicit per-column widths (w:col children);
   *  true (the default) = equal split of the content box. */
  equalWidth: boolean;
  /** Explicit column widths, px (w:col/@w:w) — present only when equalWidth
   *  is false; gaps stay `spacePx` between every pair. */
  columnsPx?: number[];
}
