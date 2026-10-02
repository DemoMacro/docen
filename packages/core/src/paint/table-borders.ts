import type { LayoutBorderEdge } from "@docen/layout";

/** One cell's own edge declarations the collapsed-border resolver reads —
 *  the shape both the flow painter and the graphic-frame painter carry. */
export interface CollapsedBorderCell {
  row: number;
  col: number;
  spanW: number;
  spanH: number;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>;
}

export interface CollapsedBorderDefaults {
  top?: LayoutBorderEdge;
  right?: LayoutBorderEdge;
  bottom?: LayoutBorderEdge;
  left?: LayoutBorderEdge;
  insideHorizontal?: LayoutBorderEdge;
  insideVertical?: LayoutBorderEdge;
}

/** One resolved stroke in table-local px: `horizontal` strokes run along a
 *  row boundary (x = grid column offset, y = grid row offset), vertical
 *  strokes the mirror. The painter offsets these by the table origin. */
export interface CollapsedBorderSegment {
  x: number;
  y: number;
  lengthPx: number;
  horizontal: boolean;
  edge: LayoutBorderEdge;
}

/** One border edge's conflict weight: an explicit nil/none erases the edge;
 *  every other edge carries its width (a style-less edge is solid). */
function edgeWeight(edge: LayoutBorderEdge | undefined): number {
  return edge && edge.style !== "nil" && edge.style !== "none" && edge.px != null ? edge.px : 0;
}

function sameEdge(a: LayoutBorderEdge, b: LayoutBorderEdge): boolean {
  return a === b || (a.px === b.px && a.style === b.style && a.color === b.color);
}

function heaviest(
  a: LayoutBorderEdge | undefined,
  b: LayoutBorderEdge | undefined,
): LayoutBorderEdge | undefined {
  return edgeWeight(b) > edgeWeight(a) ? b : a;
}

/** Collapsed borders shared by flow and graphic-frame tables: every boundary
 *  resolves to its heaviest candidate — the two adjacent cells' own edges and
 *  the table-level default (rim edges for the grid's outline, inside edges
 *  between cells). Word's conflict rule is width-first; ties keep the earlier
 *  candidate. Contiguous slots with the same winner merge into one segment. */
export function collapsedBorderSegments(
  colX: readonly number[],
  rowY: readonly number[],
  cells: readonly CollapsedBorderCell[],
  tableBorders?: CollapsedBorderDefaults,
): CollapsedBorderSegment[] {
  const nRows = rowY.length - 1;
  const nCols = colX.length - 1;
  const occ = Array.from({ length: nRows }, (_, row) =>
    Array.from({ length: nCols }, (_, col) =>
      cells.find(
        (cell) =>
          cell.row <= row &&
          row < cell.row + cell.spanH &&
          cell.col <= col &&
          col < cell.col + cell.spanW,
      ),
    ),
  );
  const tb = tableBorders;
  /** A horizontal boundary (row edge `b`) at column `c`: the cell above ends
   *  here, the cell below starts here. */
  const pickH = (b: number, c: number): LayoutBorderEdge | undefined => {
    const above = b > 0 ? occ[b - 1][c] : undefined;
    const below = b < nRows ? occ[b][c] : undefined;
    if (above && above === below) return undefined;
    // An explicitly declared cell edge (w:tcBorders, nil included) suppresses
    // the table-level default — Word resolves tcBorders over tblBorders
    // outright; only cells silent on the edge fall back to the grid default.
    const aboveEdge = above?.borders?.bottom;
    const belowEdge = below?.borders?.top;
    const def =
      aboveEdge || belowEdge
        ? undefined
        : b === 0
          ? tb?.top
          : b === nRows
            ? tb?.bottom
            : tb?.insideHorizontal;
    return heaviest(aboveEdge, heaviest(belowEdge, def));
  };
  /** A vertical boundary (column edge `b`) at row `r`. */
  const pickV = (b: number, r: number): LayoutBorderEdge | undefined => {
    const left = b > 0 ? occ[r][b - 1] : undefined;
    const right = b < nCols ? occ[r][b] : undefined;
    if (left && left === right) return undefined;
    const leftEdge = left?.borders?.right;
    const rightEdge = right?.borders?.left;
    const def =
      leftEdge || rightEdge
        ? undefined
        : b === 0
          ? tb?.left
          : b === nCols
            ? tb?.right
            : tb?.insideVertical;
    return heaviest(leftEdge, heaviest(rightEdge, def));
  };
  const out: CollapsedBorderSegment[] = [];
  for (let b = 0; b <= nRows; b++) {
    let segStart = -1;
    let seg: LayoutBorderEdge | undefined;
    for (let c = 0; c <= nCols; c++) {
      const winner = c < nCols ? pickH(b, c) : undefined;
      if (seg && winner && sameEdge(seg, winner)) continue;
      if (seg)
        out.push({
          x: colX[segStart]!,
          y: rowY[b]!,
          lengthPx: colX[c]! - colX[segStart]!,
          horizontal: true,
          edge: seg,
        });
      seg = winner && edgeWeight(winner) > 0 ? winner : undefined;
      segStart = seg ? c : -1;
    }
  }
  for (let b = 0; b <= nCols; b++) {
    let segStart = -1;
    let seg: LayoutBorderEdge | undefined;
    for (let r = 0; r <= nRows; r++) {
      const winner = r < nRows ? pickV(b, r) : undefined;
      if (seg && winner && sameEdge(seg, winner)) continue;
      if (seg)
        out.push({
          x: colX[b]!,
          y: rowY[segStart]!,
          lengthPx: rowY[r]! - rowY[segStart]!,
          horizontal: false,
          edge: seg,
        });
      seg = winner && edgeWeight(winner) > 0 ? winner : undefined;
      segStart = seg ? r : -1;
    }
  }
  return out;
}
