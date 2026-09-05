import type {
  FontMetrics,
  LaidOutParagraph,
  ProjectedColumns,
  ProjectedFlowBox,
  ProjectedLineNumbers,
  ProjectedPageBackground,
  ProjectedPageFurniture,
} from "@docen/layout";

/** One hit-testable drawing box, page-local px — what a click needs to grab a
 *  drawing (Word: clicking a picture selects it). `para` is the laid host
 *  paragraph (the caret map resolves it to the PM position) and `index` the
 *  drawing's position among that paragraph's drawings, matching the run order
 *  projectDrawings collected them in. */
export interface DrawingHitBox {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  para: LaidOutParagraph;
  index: number;
  /** "drawing" — a floating picture/shape from para.drawings (index counts
   *  that sequence); "inline" — a picture line item (index counts the
   *  paragraph's inline pictures). The PM side re-finds the node per kind. */
  kind: "drawing" | "inline";
  /** Clockwise rotation of the box about its center, degrees — the click's
   *  hit test un-rotates the point into the box's own space. */
  rotation?: number;
}

/** One line-number label on this page — content-flow yPx (the painter adds
 *  the page's content origin) plus the number and the strut size it paints
 *  at (the counted line's ¶-mark size — Word renders the numbers in the
 *  paragraph's run font size). */
export interface LineNumberMark {
  yPx: number;
  num: number;
  sizePx: number;
}

/** The paint context for one page — the stage context plus the page's own
 *  identity (page-number fields resolve against it) and which of Word's two
 *  text-underlapping layers is being composed right now (the stage paints a
 *  page twice: once for behind-doc floats, once for everything else with
 *  header/footer furniture between them — Word renders footer furniture and
 *  the body over those floats, so furniture must not sit under them).
 *
 *  The flow box and furniture are the PAGE's OWN section's (multi-section
 *  documents give every page the box of the section it belongs to). */
export interface PaintContext {
  metrics: FontMetrics;
  flow: ProjectedFlowBox;
  furniture?: ProjectedPageFurniture;
  /** This page's line numbers (w:lnNumType): the section's config plus the
   *  marks from the stage's cross-page count. Absent = the section has none. */
  lineNumbers?: { config: ProjectedLineNumbers; marks: LineNumberMark[] };
  /** The page's section columns (w:cols) — separator lines paint between
   *  them when `separate` is set. Absent = single column, nothing to draw. */
  columns?: ProjectedColumns;
  background?: ProjectedPageBackground;
  pageIndex: number;
  pageCount: number;
  layer: "behind" | "body";
  /** Forces a frame after an async image insert: Leafer's change-driven
   *  scheduling stalls on apps created while offscreen (see stage.repaint),
   *  so a decode completing after repaint would otherwise never show. */
  rerender: () => void;
  /** Accumulates this page's drawing boxes as the body pass paints them —
   *  the stage turns the list into its click hit table. */
  hitBoxes?: DrawingHitBox[];
  /** In-front floats park here instead of painting inside their anchor
   *  paragraph: Word stacks them above ALL text (an anchor earlier in the
   *  flow must not let later paragraphs paint over the float), so the stage
   *  flushes this queue after the body pass paints its last paragraph.
   *  Behind-doc floats defer too: the stage sorts each band by
   *  w:relativeHeight before executing, so same-band stacking follows the
   *  document's z-order instead of document order. The entry carries its
   *  band — the flush restores ctx.layer per band so member painting keeps
   *  the layer semantics it had with direct painting. */
  deferredDrawings?: Array<{ z: number; layer: "behind" | "body"; paint: () => void }>;
  /** Formatting marks visibility (Word's ¶ toggle): the painter draws the
   *  bent arrow at line/paragraph ends, an arrow on tabs, a dot on spaces,
   *  and the break rows (page/section) between blocks. */
  showMarks?: boolean;
  /** Document-grid overlay (Word's View → Gridlines): horizontal rules every
   *  `linePitchPx` across the content box, painted under the body text. */
  showGridlines?: boolean;
  /** The break rows' labels (Word paints them in the UI language). The
   *  section-break one has per-type variants — Word names the actual break
   *  ("分节符(连续)"), defaulting to the next-page label. */
  marksLabels?: {
    pageBreak?: string;
    sectionBreak?: string;
    sectionBreakContinuous?: string;
    sectionBreakEvenPage?: string;
    sectionBreakOddPage?: string;
  };
}

/** The text column a block paints inside: the page's content box for body
 *  blocks, the cell's inner box for table content (a text box's insets box
 *  for its paragraphs). Cell-anchored floats clamp inside it — Word's
 *  layoutInCell containment, matching the wrap zones the layout built. */
export interface PaintColumn {
  width: number;
  inCell: boolean;
}
