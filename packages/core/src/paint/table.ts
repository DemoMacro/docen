import { tableGridOf, type LaidOutTable, type LayoutBorderEdge } from "@docen/layout";
import { Line, Rect, type IGroup } from "leafer-ui";

import { paintBlock } from "../painter";
import type { PaintContext } from "./context";
import {
  collapsedBorderSegments,
  type CollapsedBorderCell,
  type CollapsedBorderDefaults,
} from "./table-borders";

export function paintTable(
  tree: IGroup,
  table: LaidOutTable,
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  // w:jc: the whole grid (borders included) shifts as one box.
  x += table.offsetXPx ?? 0;
  // The shared walk: boundaries and every cell's content origin — the caret
  // map consumes the same tableGridOf output.
  const { colX, rowY, cells } = tableGridOf(table);

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

  drawCollapsedTableBorders(tree, x, y, colX, rowY, cells, table.borders);
}

/** Collapsed borders shared by flow and graphic-frame tables: every boundary
 *  resolves to its heaviest candidate — the two adjacent cells' own edges and
 *  the table-level default (rim edges for the grid's outline, inside edges
 *  between cells). Word's conflict rule is width-first; ties keep the earlier
 *  candidate. Contiguous slots with the same winner merge into one stroke. */
export function drawCollapsedTableBorders(
  tree: IGroup,
  x: number,
  y: number,
  colX: readonly number[],
  rowY: readonly number[],
  cells: readonly CollapsedBorderCell[],
  tableBorders?: CollapsedBorderDefaults,
): void {
  for (const segment of collapsedBorderSegments(colX, rowY, cells, tableBorders)) {
    drawEdge(
      tree,
      x + segment.x,
      y + segment.y,
      segment.lengthPx,
      segment.horizontal,
      segment.edge,
    );
  }
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
 *  parallel strokes. Shared with the graphic-frame table painter (the same
 *  visual language — hairline lift, dash map, multi-stroke composites). */
export function drawEdge(
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
