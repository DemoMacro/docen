import { tableGridOf, type LaidOutTable, type LayoutBorderEdge } from "@docen/layout";
import { Line, Rect, type IGroup } from "leafer-ui";

import { paintBlock } from "../painter";
import type { PaintContext } from "./context";

export function paintTable(
  tree: IGroup,
  table: LaidOutTable,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  // w:jc: the whole grid (borders included) shifts as one box.
  x += table.offsetXPx ?? 0;
  // The shared walk: boundaries, the occupancy grid, and every cell's
  // content origin — the caret map consumes the same tableGridOf output.
  const { colX, rowY, occ, cells } = tableGridOf(table);
  const nRows = table.rows.length;
  const nCols = table.columnWidthsPx.length;

  for (const p of cells) {
    // Shading covers the merged box; content anchors to the start row (the
    // engine measured it there).
    if (p.cell.fill) {
      tree.add(
        new Rect({
          x: x + colX[p.col]!,
          y: y + rowY[p.row]!,
          width: colX[p.col + p.spanW]! - colX[p.col]!,
          height: rowY[p.row + p.spanH]! - rowY[p.row]!,
          fill: `#${p.cell.fill}`,
        }),
      );
    }
    const contentX = x + p.contentXPx;
    const contentY = y + p.contentYPx;
    for (const stacked of p.cell.stack) {
      paintBlock(tree, stacked.block, contentX, contentY + stacked.yPx, ctx, {
        width: p.cell.innerWidthPx,
        inCell: true,
      });
    }
  }

  // Collapsed borders: every shared boundary resolves to its heaviest
  // candidate — the two adjacent cells' own edges and the table-level default
  // (rim edges for the grid's outline, inside edges between cells). Word's
  // conflict rule is width-first; ties keep the earlier candidate.
  const tb = table.borders;
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
  // Contiguous boundary slots with an identical winner merge into one stroke.
  for (let b = 0; b <= nRows; b++) {
    let segStart = -1;
    let seg: LayoutBorderEdge | undefined;
    for (let c = 0; c <= nCols; c++) {
      const winner = c < nCols ? pickH(b, c) : undefined;
      if (seg && winner && sameEdge(seg, winner)) continue;
      if (seg) {
        drawEdge(tree, x + colX[segStart], y + rowY[b], colX[c] - colX[segStart], true, seg);
      }
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
      if (seg) {
        drawEdge(tree, x + colX[b], y + rowY[segStart], rowY[r] - rowY[segStart], false, seg);
      }
      seg = winner && edgeWeight(winner) > 0 ? winner : undefined;
      segStart = seg ? r : -1;
    }
  }
}

/** One border edge's conflict weight: nil/none/absent carry none. */
function edgeWeight(edge: LayoutBorderEdge | undefined): number {
  return edge && edge.style && edge.style !== "nil" && edge.style !== "none" && edge.px != null
    ? edge.px
    : 0;
}

/** Word's border conflict resolution: the wider edge wins; ties keep `a`. */
function heaviest(
  a: LayoutBorderEdge | undefined,
  b: LayoutBorderEdge | undefined,
): LayoutBorderEdge | undefined {
  return edgeWeight(b) > edgeWeight(a) ? b : a;
}

function sameEdge(a: LayoutBorderEdge, b: LayoutBorderEdge): boolean {
  return a === b || (a.px === b.px && a.style === b.style && a.color === b.color);
}

/** dashPattern per OOXML border style (stroke-only in Leafer); styles without
 *  an entry render solid — the visual fallback for wave/3D composites. */
const DASH_PATTERN: Record<string, number[]> = {
  dashed: [4, 2],
  dashSmallGap: [2, 2],
  dotted: [1, 2],
  dashDot: [4, 2, 1, 2],
  dashDotDot: [4, 2, 1, 2, 1, 2],
};

/** Draw one collapsed edge centered on its boundary: a stroked Line (dash
 *  styles apply to strokes, not fills); double/triple split the width into
 *  parallel strokes. */
function drawEdge(
  tree: IGroup,
  ex: number,
  ey: number,
  len: number,
  horizontal: boolean,
  edge: LayoutBorderEdge,
): void {
  // Word's screen rendering lifts hairlines to a full pixel at 100% zoom —
  // a sub-pixel stroke here would render as a faint half-transparent line.
  const px = Math.max(edge.px ?? 1, 1);
  const color = edge.color ? `#${edge.color}` : "#000000";
  const style = edge.style ?? "single";
  const dash = DASH_PATTERN[style];
  const stroke = (offset: number, thickness: number): void => {
    const wx = horizontal ? ex : ex + offset;
    const wy = horizontal ? ey + offset : ey;
    tree.add(
      new Line({
        x: wx,
        y: wy,
        // Line points are relative to x/y; a zero-length second point pins the
        // direction (horizontal → +x, vertical → +y).
        points: horizontal ? [0, 0, len, 0] : [0, 0, 0, len],
        stroke: color,
        strokeWidth: thickness,
        dashPattern: dash,
      }),
    );
  };
  if (style === "double" || style === "triple") {
    const unit = px / (style === "double" ? 3 : 4);
    stroke(-px / 2 + unit / 2, unit);
    if (style === "triple") stroke(0, unit);
    stroke(px / 2 - unit / 2, unit);
    return;
  }
  stroke(0, px);
}
