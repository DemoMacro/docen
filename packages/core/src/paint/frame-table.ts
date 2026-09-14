import {
  stackBlocks,
  TextMeasurer,
  type LaidOutStackItem,
  type LayoutBlock,
  type LayoutBorderEdge,
} from "@docen/layout";
import { Rect, type IGroup } from "leafer-ui";

import { paintBlock } from "../painter";
import type { PaintContext } from "./context";
import { drawEdge } from "./table";

// ── graphic-frame table painter ──
//
// Draws a `kind: "table"` member (a pptx a:tbl) from the projection's
// normalized payload — px geometry, origin cells with resolved spans, and
// laid text blocks (consumed structurally, like the chart payload; this
// package holds no office-open types). PowerPoint's model is explicit
// geometry: column widths are fixed, a row height is a minimum the content
// grows, and every cell paints its own declared edges (no Word-style border
// collapse — the later cell's edge sits on top).

interface FrameCell {
  col: number;
  row: number;
  spanW: number;
  spanH: number;
  fill?: string;
  opacity?: number;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>;
  anchor?: "top" | "center" | "bottom";
  marginsPx: { left: number; top: number; right: number; bottom: number };
  blocks: LayoutBlock[];
}

interface FrameRow {
  heightPx: number;
  cells: FrameCell[];
}

/** The payload shape the projection builds (see the pptx tableMember). */
interface FrameTablePayload {
  columnWidthsPx: number[];
  rows: FrameRow[];
}

export function paintFrameTable(
  tree: IGroup,
  member: { x: number; y: number; table: unknown },
  ctx: PaintContext,
): void {
  const t = member.table as FrameTablePayload | undefined;
  const cols = t?.columnWidthsPx;
  const rows = t?.rows;
  if (!cols?.length || !rows?.length) return;
  const measurer = new TextMeasurer(ctx.metrics);

  // Column bands from the fixed widths.
  const colX = [0];
  for (const w of cols) colX.push(colX[colX.length - 1]! + w);

  // Measure every cell's content once; a row grows to its tallest cell, and
  // a vertically merged cell's need lands on the last row of its span (the
  // rows above keep their declared height). The same stack paints below, so
  // the height the row grew by and the drawn ink agree.
  const laid: {
    cell: FrameCell;
    innerWidth: number;
    stack: LaidOutStackItem[];
    heightPx: number;
  }[] = [];
  const rowNeed = rows.map(() => 0);
  for (const row of rows) {
    for (const cell of row.cells) {
      const innerWidth = Math.max(
        1,
        (colX[cell.col + cell.spanW] ?? colX[colX.length - 1]!) -
          colX[cell.col]! -
          cell.marginsPx.left -
          cell.marginsPx.right,
      );
      const stacked = stackBlocks(cell.blocks, innerWidth, undefined, measurer);
      const need = stacked.heightPx + cell.marginsPx.top + cell.marginsPx.bottom;
      const last = Math.min(cell.row + cell.spanH - 1, rows.length - 1);
      let declaredAbove = 0;
      for (let r = cell.row; r < last; r++) declaredAbove += rows[r]?.heightPx ?? 0;
      rowNeed[last] = Math.max(rowNeed[last]!, need - declaredAbove);
      laid.push({ cell, innerWidth, stack: stacked.stack, heightPx: stacked.heightPx });
    }
  }
  const rowH = rows.map((row, r) => Math.max(row.heightPx, rowNeed[r]!));
  const rowY = [0];
  for (const h of rowH) rowY.push(rowY[rowY.length - 1]! + h);

  const spanEnd = (v: number[], i: number, n: number): number => v[Math.min(i + n, v.length - 1)]!;

  // Fills first, then text, then every cell's own edges on top (PowerPoint's
  // paint order — a later cell's border sits over its neighbor's fill).
  for (const { cell } of laid) {
    if (!cell.fill) continue;
    tree.add(
      new Rect({
        x: member.x + colX[cell.col]!,
        y: member.y + rowY[cell.row]!,
        width: spanEnd(colX, cell.col, cell.spanW) - colX[cell.col]!,
        height: spanEnd(rowY, cell.row, cell.spanH) - rowY[cell.row]!,
        fill: `#${cell.fill}`,
        // A faded fill rides the element's own opacity — the cell Rect
        // carries no stroke of its own to keep opaque.
        ...(cell.opacity != null ? { opacity: cell.opacity } : {}),
      }),
    );
  }
  for (const { cell, innerWidth, stack, heightPx } of laid) {
    const bandHeight = spanEnd(rowY, cell.row, cell.spanH) - rowY[cell.row]!;
    const slack = bandHeight - cell.marginsPx.top - cell.marginsPx.bottom - heightPx;
    const lead =
      slack > 0 ? (cell.anchor === "bottom" ? slack : cell.anchor === "center" ? slack / 2 : 0) : 0;
    const contentY = member.y + rowY[cell.row]! + cell.marginsPx.top + lead;
    for (const item of stack) {
      paintBlock(
        tree,
        item.block,
        member.x + colX[cell.col]! + cell.marginsPx.left,
        contentY + item.yPx,
        ctx,
        { width: innerWidth, inCell: true },
      );
    }
  }
  for (const { cell } of laid) {
    if (!cell.borders) continue;
    const x0 = member.x + colX[cell.col]!;
    const y0 = member.y + rowY[cell.row]!;
    const w = spanEnd(colX, cell.col, cell.spanW) - colX[cell.col]!;
    const h = spanEnd(rowY, cell.row, cell.spanH) - rowY[cell.row]!;
    if (cell.borders.top) drawEdge(tree, x0, y0, w, true, cell.borders.top);
    if (cell.borders.bottom) drawEdge(tree, x0, y0 + h, w, true, cell.borders.bottom);
    if (cell.borders.left) drawEdge(tree, x0, y0, h, false, cell.borders.left);
    if (cell.borders.right) drawEdge(tree, x0 + w, y0, h, false, cell.borders.right);
  }
}
