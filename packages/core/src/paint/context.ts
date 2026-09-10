import type {
  FontMetrics,
  LaidOutBlock,
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
  /** A group member's index path through the wpgGroup's PM content (absent =
   *  the drawing's own box, or a rotated drawing's member whose geometry the
   *  hit test cannot un-map). Painted after the group box, so the member
   *  wins the click; whether it selects the member or the group is the
   *  editor's state call (Word: a click selects the group until it is
   *  entered). */
  childPath?: readonly number[];
  /** A behind-doc float's box — only the in-front band covers the text layer,
   *  so overlays that must yield to front floats (spelling squiggles) skip
   *  these boxes. */
  behind?: boolean;
  /** Clockwise rotation of the box about its center, degrees — the click's
   *  hit test un-rotates the point into the box's own space. */
  rotation?: number;
  /** A chart sub-element the box selects inside the framed chart (Word's
   *  two-stage: the first click selects the chart, the next lands on the
   *  plot). Present only on boxes the chart painter registered. */
  chartPart?: ChartPartHit;
}

/** One chart sub-element's exact geometry when it is not a rectangle (a pie
 *  wedge, a series line) — page-local px; the hit box's rect stays the
 *  bounding box. A wedge is the sector between the two angles (radii `r`
 *  outer, `hole` inner); a poly is the point run, hit-tested as a corridor
 *  of `width` around the segments or, when `closed`, as the polygon. */
export type ChartPartShape =
  | { kind: "wedge"; cx: number; cy: number; r: number; a0: number; a1: number; hole?: number }
  | { kind: "poly"; pts: [number, number][]; width?: number; closed?: boolean };

/** A chart sub-element a click can select inside the framed chart.
 *  `series`/`point` index the chart payload's series array and a series'
 *  values; `legend` marks a legend entry (it selects its series, like
 *  Word's); `title` the title band. `valueDrag` marks the element as
 *  draggable to change its value (Excel): a linear page-local px → value
 *  map, `value = a + b·px`, reading the pointer's y on vertical charts
 *  and x on horizontal ones (`horizontal`). Present on clustered/stacked
 *  bars and line/area points; absent where a drag has no meaning (pie
 *  wedges, percentStacked shares, scatter). */
export interface ChartPartHit {
  series?: number;
  point?: number;
  legend?: boolean;
  title?: boolean;
  shape?: ChartPartShape;
  valueDrag?: { a: number; b: number; horizontal?: boolean };
}

/** The exact-shape hit test behind {@link ChartPartShape}. */
export function chartShapeHit(shape: ChartPartShape, x: number, y: number): boolean {
  if (shape.kind === "wedge") {
    const dx = x - shape.cx;
    const dy = y - shape.cy;
    const d = Math.hypot(dx, dy);
    if (d > shape.r || (shape.hole != null && d < shape.hole)) return false;
    if (shape.a1 - shape.a0 >= Math.PI * 2) return true;
    let a = Math.atan2(dy, dx);
    const twoPi = Math.PI * 2;
    while (a < shape.a0) a += twoPi;
    return a <= shape.a1;
  }
  const pts = shape.pts;
  if (!shape.closed) {
    const half = (shape.width ?? 10) / 2;
    for (let i = 1; i < pts.length; i++) {
      const [x0, y0] = pts[i - 1]!;
      const [x1, y1] = pts[i]!;
      const dx = x1 - x0;
      const dy = y1 - y0;
      const t =
        dx === 0 && dy === 0
          ? 0
          : Math.max(0, Math.min(1, ((x - x0) * dx + (y - y0) * dy) / (dx * dx + dy * dy)));
      if (Math.hypot(x - (x0 + t * dx), y - (y0 + t * dy)) <= half) return true;
    }
    return false;
  }
  let inside = false;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const [xi, yi] = pts[i]!;
    const [xj, yj] = pts[j]!;
    if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

/** What the chart painter needs to register sub-element hit boxes — the
 *  paint context (page, layer, the accumulation list) plus the host identity
 *  a click resolves through. Absent for rotated/mirrored drawings: their
 *  members paint in spinner space whose geometry the hit test cannot
 *  un-map. */
export interface ChartHitContext {
  ctx: PaintContext;
  para: LaidOutParagraph;
  index: number;
  kind: DrawingHitBox["kind"];
  /** The member space's page origin. A holder Box (inline pictures, clipped
   *  drawings) positions members tree-local while sitting at (ox,oy) on the
   *  page, so registered boxes add it back; direct painters carry the origin
   *  in the member position and pass 0. */
  ox: number;
  oy: number;
}

/** One editable text-box stack (a wps txbx member's laid paragraphs),
 *  page-local px — the caret map registers its lines against the shape's PM
 *  content so a double click edits the text in place. `host` re-finds the
 *  wpsShape node (the same identity a drawing hit box carries; `childPath`
 *  drills into a group's interior). Metafile text (drawn GDI art, `nowrap`)
 *  and rotated stacks paint but never register. */
export interface ShapeTextStack {
  page: number;
  host: { para: LaidOutParagraph; index: number; childPath?: readonly number[] };
  /** The insets box origin the blocks stack from. */
  xPx: number;
  yPx: number;
  /** The laid blocks as stackBlocks returned them. */
  items: readonly { yPx: number; block: LaidOutBlock }[];
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
  /** Accumulates this page's editable text-box stacks the same way — the
   *  bridge registers them with the caret map (double-click-to-edit). */
  shapeTextStacks?: ShapeTextStack[];
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
  /** The stack renders inside a drawing shape (a wps txbx / metafile
   *  text-box member): its lines follow Word's DrawingML text-box baseline
   *  model — the element's own 0.85 × size share IS the baseline depth —
   *  not the body's measured-ascent model. Metafile strings additionally
   *  encode the GDI baseline in each line top (the −0.8 × size
   *  calibration), which the same share absorbs. Pixel-verified against
   *  the reference PDF: the header banner slogan rides the 0.85 anchor. */
  shapeText?: boolean;
}
