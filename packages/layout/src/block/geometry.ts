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

/** A line's x origin relative to its block: the left indent (every line),
 *  the line's own first-line indent flag (a split tail carries none), and a
 *  wrapSide float's shift — the exact sum the painter offsets by. */
export function lineOriginXPx(para: LaidOutParagraph, line: LaidOutLine): number {
  return (para.indent?.leftPx ?? 0) + (line.firstLineIndentPx ?? 0) + (line.xOffsetPx ?? 0);
}

/** A docGrid body line's half-leading: Word centers the run's EM box in the
 *  grid span (the browser font box the natural height measures runs deeper —
 *  corpus-verified on the honor table, ~0.3em of it); every other regime
 *  anchors at the line top (pad 0). Text-box stacks take the same rule —
 *  their grid-snapped lines half-lead like the body's (pixel-verified: the
 *  reference render's first ink sits at half-leading in a box whose border
 *  position matches ours exactly), and bodyPr @compatLnSpc changes nothing
 *  (Word ignores it when laying out wps txbxContent). Both the painter's
 *  text y and the caret band anchor at this pad. */
export function gridPadOf(line: LaidOutLine): number {
  if (!line.grid) return 0;
  // A picture-floored line centers the picture box (its natural) in the
  // spanned rows — beside-text pictures must not inherit the text EM ref.
  const ref = line.pictureFloored ? line.naturalPx : (line.textEmPx ?? line.naturalPx);
  return Math.max(0, (line.heightPx - ref) / 2);
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
  const ends: number[] = new Array(line.items.length);
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
