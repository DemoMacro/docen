// Shared layout geometry — the derived line/table math both consumers of a
// laid-out page need: the scene painter (drawing) and the editor's caret map
// (hit-testing). One implementation here means the caret can never drift
// from the paint — each function is the single authority for its sum.

import type {
  LaidOutBlock,
  LaidOutCell,
  LaidOutLine,
  LaidOutParagraph,
  LaidOutTable,
} from "../layout-result";
import { leaferBaselinePadPx } from "../text/measure";

/** A line's x origin relative to its block: the left indent (every line),
 *  the line's own first-line indent flag (a split tail carries none), and a
 *  wrapSide float's shift — the exact sum the painter offsets by. */
export function lineOriginXPx(para: LaidOutParagraph, line: LaidOutLine): number {
  return (para.indent?.leftPx ?? 0) + (line.firstLineIndentPx ?? 0) + (line.xOffsetPx ?? 0);
}

/** A line's leading pad above its text. A docGrid body line centers the run's
 *  EM box in the grid span (Word's half-leading — the browser font box the
 *  natural height measures runs deeper — corpus-verified on the honor table,
 *  ~0.3em of it). A non-grid multiple-spacing line scales the WHOLE box
 *  (Word): the baseline rides at factor × its single-line height, so the
 *  slack splits on the natural box's own ascent:descent ratio instead of
 *  sinking below the text — the single-line baseline height is the paint
 *  anchor (leaferBaselinePadPx). A picture-floored line keeps its slack
 *  below the box: an inline picture sits ON the baseline, whose single-line
 *  position IS the box bottom, so scaling leaves the box top-anchored.
 *  atLeast lines stay top-anchored (Word rides the extra space BELOW the
 *  text). exact lines sink onto the box bottom — the slack piles ABOVE the
 *  text — but never rise past the top: Word's undersized exact keeps the
 *  glyphs at their natural position (clipping their tops at the box edge,
 *  pixel-verified) rather than lifting them, so the pad clamps at 0.
 *  Text-box stacks take the grid rule — their grid-snapped lines half-lead
 *  like the body's (pixel-verified: the reference render's first ink sits at
 *  half-leading in a box whose border position matches ours exactly), and
 *  bodyPr @compatLnSpc changes nothing (Word ignores it when laying out wps
 *  txbxContent). Both the painter's text y and the caret band anchor at this
 *  pad. */
export function gridPadOf(line: LaidOutLine): number {
  if (line.spacingRule === "exact") return Math.max(0, line.heightPx - line.naturalPx);
  if (line.grid) {
    // A picture-floored line centers the picture box (its natural) in the
    // spanned rows — beside-text pictures must not inherit the text EM ref.
    const ref = line.pictureFloored ? line.naturalPx : (line.textEmPx ?? line.naturalPx);
    return Math.max(0, (line.heightPx - ref) / 2);
  }
  if (line.pictureFloored || line.spacingRule === "atLeast") {
    return 0;
  }
  if (!line.textEmPx || line.naturalPx <= 0) return 0;
  const slack = line.heightPx - line.naturalPx;
  if (slack <= 0) return 0;
  // The single-line baseline height the multiple scales: the measured face
  // ascent (baselinePadPx), falling back to Leafer's 0.85 constant on
  // fixtures that predate the field.
  const singleBaseline = line.baselinePadPx ?? leaferBaselinePadPx(line.textEmPx);
  return (slack * singleBaseline) / line.naturalPx;
}

/** A shape-text line's leading pad: only the container pads — the grid
 *  half-lead and the exact-line sink, the two pixel-verified inside text
 *  boxes. The body's multiple-spacing scale stays a flow behavior: a
 *  DrawingML body anchors its first baseline at the font box (inset +
 *  ascent), and the stack's own line heights already space later lines, so
 *  the scale's extra leads BELOW each line — applying it above sinks a
 *  box's first line out of its artwork (pixel-verified: the header banner
 *  slogan inheriting the document default's double spacing). */
export function shapeTextPadPx(line: LaidOutLine): number {
  return line.grid || line.spacingRule === "exact" ? gridPadOf(line) : 0;
}

/** A line's alphabetic baseline depth below its top (px): the leading pad
 *  (gridPadOf) plus the dominant face's measured ascent (the layout's
 *  per-line `baselinePadPx`), falling back to Leafer's 0.85 × size constant
 *  on lines without the field (`fallbackSizePx` covers textless lines, where
 *  the paragraph-mark strut is the only baseline reference). A picture-floored
 *  line pushes the baseline to the picture's bottom — Word treats an inline
 *  graphic as a single character whose bottom edge sits ON the baseline
 *  ("images are aligned to the baseline", so taller pictures sink the line's
 *  baseline and shorter pictures bottom-align with them). The ONE anchor
 *  every text consumer hangs off — the painter's glyphs, underline, tab
 *  leaders and formatting marks, plus the editor's caret band and line
 *  numbers all position relative to this, so none can drift from the paint. */
export function lineBaselineDepthPx(line: LaidOutLine, fallbackSizePx = 0): number {
  const pad = gridPadOf(line);
  const textDepth =
    pad + (line.baselinePadPx ?? leaferBaselinePadPx(line.textEmPx ?? fallbackSizePx));
  // The floored natural IS the tallest picture (line-break floors it), so the
  // picture's bottom = pad + naturalPx; a text-only line keeps the text depth.
  const pictureDepth = line.pictureFloored ? pad + line.naturalPx : 0;
  return Math.max(textDepth, pictureDepth);
}

/** A block's page-fitting extent — its content bottom. A paragraph whose
 *  last line is a picture spans grid rows whose trailing half-leading may
 *  overhang the page bottom: Word keeps the picture when its own box fits
 *  (pixel-verified — a 639px picture with 9px leading stays on a page 6px
 *  short of the padded box). The flow's cursor still advances by the full
 *  padded box; only the fit check sees this extent. */
export function fitExtentPx(block: LaidOutBlock): number {
  if (block.kind !== "paragraph") return block.heightPx;
  const last = block.lines[block.lines.length - 1];
  if (!last?.pictureFloored) return block.heightPx;
  return block.heightPx - last.heightPx + gridPadOf(last) + last.naturalPx;
}

/** Each item's justified stretch-interval end — a text item's glyphs fill
 *  from its own x to the next item's post-justify x, the last item to
 *  maxWidth + the overflow-punct hang. Null on unjustified lines. */
export function justifiedIntervals(line: LaidOutLine): number[] | null {
  if (line.justifyGapPx == null) return null;
  const ends: number[] = Array.from({ length: line.items.length });
  let nextLeft = (line.maxWidthPx ?? 0) + (line.hangPx ?? 0);
  for (let i = line.items.length - 1; i >= 0; i--) {
    ends[i] = nextLeft;
    nextLeft = line.items[i]!.xPx;
  }
  return ends;
}

/** One placed cell of the shared table walk. */
export interface TableCellPlacement {
  cell: LaidOutCell;
  /** Start grid column / row of the cell's span. */
  col: number;
  row: number;
  spanW: number;
  spanH: number;
  /** Table-relative content origin: the spanned column's left edge + the
   *  cell insets, the start row's top + insets + the vertical-align offset. */
  contentXPx: number;
  contentYPx: number;
}

export interface TableGrid {
  /** Column left edges + the right rim (nCols + 1 entries); row tops + the
   *  bottom rim. */
  colX: number[];
  rowY: number[];
  /** occ[r][c] = the cell covering that grid slot (spanned slots included) —
   *  boundary resolution sees across a span and skips its inner edges. */
  occ: (LaidOutCell | undefined)[][];
  cells: TableCellPlacement[];
}

/** The table walk both the painter (shading, borders, content) and the caret
 *  map (paragraph anchoring) need: boundary coordinates, the occupancy grid,
 *  and every cell's content origin. One traversal, so hit-testing anchors
 *  content exactly where painting does. */
export function tableGridOf(table: LaidOutTable): TableGrid {
  const colX = [0];
  for (const w of table.columnWidthsPx) colX.push(colX[colX.length - 1] + w);
  const rowY = [0];
  for (const row of table.rows) rowY.push(rowY[rowY.length - 1] + row.heightPx);
  const nCols = table.columnWidthsPx.length;
  const nRows = table.rows.length;
  const occ: (LaidOutCell | undefined)[][] = Array.from({ length: nRows }, () =>
    Array.from<LaidOutCell | undefined>({ length: nCols }),
  );
  const cells: TableCellPlacement[] = [];
  table.rows.forEach((row, r) => {
    let col = 0;
    for (const cell of row.cells) {
      while (col < nCols && occ[r]![col]) col++;
      if (col >= nCols) break;
      const spanW = Math.min(cell.colspan, nCols - col);
      const spanH = Math.min(cell.rowspan ?? 1, nRows - r);
      for (let dr = 0; dr < spanH; dr++)
        for (let dc = 0; dc < spanW; dc++) occ[r + dr]![col + dc] = cell;
      cells.push({
        cell,
        col,
        row: r,
        spanW,
        spanH,
        contentXPx: colX[col]! + (cell.insets.left ?? 0),
        contentYPx: rowY[r]! + (cell.insets.top ?? 0) + (cell.contentOffsetYPx ?? 0),
      });
      col += spanW;
    }
  });
  return { colX, rowY, occ, cells };
}
