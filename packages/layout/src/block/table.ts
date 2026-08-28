// Table layout — w:tbl ≡ a:tbl share this geometry. Ported from the editor's
// measureRowHeight: a row is its tallest cell; a cell's content stacks with
// collapsing paragraph margins inside its insets; trHeight floors (atLeast)
// or fixes (exact) the row.
//
// Known boundary (inherited from measure.ts, deliberate for P1): a rowspan
// cell's full content counts on its START row — Word distributes across the
// span. Start-row over-estimate + spanned-row under-estimate is deterministic
// (no re-flow wobble); mid-row page splitting lands with flow/ in P2.

import type {
  LayoutBlockContext,
  LayoutBorderEdge,
  LayoutCellInsets,
  LayoutTable,
} from "../layout-doc";
import type { LaidOutCell, LaidOutTable } from "../layout-result";
import type { TextMeasurer } from "../text/measure";
import { stackBlocks } from "./block";

/** Lay out a table at its container's content width. */
export function layoutTable(
  table: LayoutTable,
  containerWidth: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): LaidOutTable {
  const widthPx = tableWidthOf(table, containerWidth);
  const columnWidths = tableColumnWidths(table, widthPx);

  let heightPx = 0;
  const rows = table.rows.map((row, rowIndex) => {
    // Cells measure under the table-cell line-height rule (max(natural,
    // pitch)) so trHeight governs. Zones start empty per cell — the cell's
    // stacker accumulates its own floats' zones (never the page's: a cell's
    // width is its column).
    const cellCtx: LayoutBlockContext = {
      ...ctx,
      inTable: true,
      floatZones: undefined,
      startY: undefined,
    };
    let rowHeight = 0;
    let colCursor = 0;
    // (cell total height, for the vAlign slack pass after the row height settles)
    const totals: number[] = [];
    const cells: LaidOutCell[] = row.cells.map((cell) => {
      const colspan = cell.colspan ?? 1;
      let cellWidth = 0;
      for (let i = 0; i < colspan && colCursor + i < columnWidths.length; i++) {
        cellWidth += columnWidths[colCursor + i];
      }
      // A cell's missing border side falls back to the table-level default —
      // the grid rim (first/last row/column) takes the outer edges, interior
      // boundaries the inside ones (CT_TblBorders semantics).
      const firstCol = colCursor === 0;
      const lastCol = colCursor + colspan >= columnWidths.length;
      colCursor += colspan;

      const insets = mergeInsets(cell.insets, table.cellInsets);
      const hInsetPx = (insets.left ?? 0) + (insets.right ?? 0);
      // Under border-collapse the column width is the cell's BORDER box, so
      // text wraps at cellWidth − insets − side borders.
      const leftEdge =
        cell.borders?.left ?? (firstCol ? table.borders?.left : table.borders?.insideVertical);
      const rightEdge =
        cell.borders?.right ?? (lastCol ? table.borders?.right : table.borders?.insideVertical);
      const topEdge =
        cell.borders?.top ??
        (rowIndex === 0 ? table.borders?.top : table.borders?.insideHorizontal);
      const bottomEdge =
        cell.borders?.bottom ??
        (rowIndex === table.rows.length - 1
          ? table.borders?.bottom
          : table.borders?.insideHorizontal);
      const innerWidthPx = Math.max(
        0,
        cellWidth - hInsetPx - borderEdgePx(leftEdge) - borderEdgePx(rightEdge),
      );

      const stacked = stackBlocks(cell.blocks, innerWidthPx, cellCtx, measurer);
      // Vertical overhead: top+bottom insets, plus only the MAX of the
      // top/bottom borders (adjacent rows share one line under collapse).
      const vOverheadPx =
        (insets.top ?? 0) +
        (insets.bottom ?? 0) +
        Math.max(borderEdgePx(topEdge), borderEdgePx(bottomEdge));

      const cellHeightPx = stacked.heightPx + vOverheadPx;
      totals.push(cellHeightPx);
      if (cellHeightPx > rowHeight) rowHeight = cellHeightPx;
      return {
        colspan,
        rowspan: cell.rowspan ?? 1,
        insets,
        borders: cell.borders,
        fill: cell.fill,
        innerWidthPx,
        stack: stacked.stack,
      };
    });

    const tr = row.height;
    if (tr && tr.px > 0) {
      rowHeight = tr.rule === "exact" ? tr.px : Math.max(rowHeight, tr.px);
    }
    // w:vAlign: place the content in the row's slack (read back from the
    // source cells — the laid-out cell carries only the resolved offset). A
    // vertically merged cell keeps its content on the start row (the
    // span-distribution boundary above), so only single-row cells shift.
    cells.forEach((cell, i) => {
      const slack = rowHeight - (totals[i] ?? 0);
      const va = row.cells[i]?.verticalAlign;
      if (slack <= 0 || (cell.rowspan ?? 1) > 1) return;
      cell.contentOffsetYPx = va === "center" ? slack / 2 : va === "bottom" ? slack : undefined;
    });
    heightPx += rowHeight;
    return { heightPx: rowHeight, cells };
  });

  return {
    kind: "table",
    widthPx,
    columnWidthsPx: columnWidths,
    // w:jc: the table box's placement in the flow column — center/right
    // against the column width, which goes negative for a table wider than
    // the column (Word centers those into the margins).
    offsetXPx:
      table.align === "center"
        ? (containerWidth - widthPx) / 2
        : table.align === "right"
          ? containerWidth - widthPx
          : undefined,
    heightPx,
    borders: table.borders,
    rows,
  };
}

/** Table content width: pct → percent of the container; px → as-is; absent →
 *  auto (fill the text column). */
function tableWidthOf(table: LayoutTable, containerWidth: number): number {
  if (!table.width) return containerWidth;
  if (table.width.type === "percent") return (containerWidth * table.width.percent) / 100;
  return table.width.px;
}

/** Per-column content widths, scaled proportionally to the effective table
 *  width (Word scales the grid to tblW, never to the raw sum). The grid comes
 *  from tblGrid, else the first row's cell widths — the grid is preferred
 *  because it is identical across page-split slices of the same table, so
 *  column widths (and row heights) stay stable across re-flows. */
function tableColumnWidths(table: LayoutTable, tableWidth: number): number[] {
  const grid: number[] = [];
  if (table.columnWidthsPx && table.columnWidthsPx.length > 0) {
    grid.push(...table.columnWidthsPx);
  } else {
    const firstRow = table.rows[0];
    if (firstRow) {
      for (const cell of firstRow.cells) {
        const colspan = cell.colspan ?? 1;
        if (cell.widthPx != null && cell.widthPx > 0) {
          for (let i = 0; i < colspan; i++) grid.push(cell.widthPx / colspan);
        } else {
          for (let i = 0; i < colspan; i++) grid.push(0);
        }
      }
    }
  }
  if (grid.length === 0) return [];
  const total = grid.reduce((a, b) => a + b, 0) || 1;
  return grid.map((w) => (w / total) * tableWidth);
}

/** A cell's own inset wins per side, else the table's default. */
function mergeInsets(
  cell: LayoutCellInsets | undefined,
  tableInsets: LayoutCellInsets | undefined,
): LayoutCellInsets {
  if (!tableInsets) return cell ?? {};
  if (!cell) return tableInsets;
  return {
    top: cell.top ?? tableInsets.top,
    right: cell.right ?? tableInsets.right,
    bottom: cell.bottom ?? tableInsets.bottom,
    left: cell.left ?? tableInsets.left,
  };
}

/** One border edge's width; nil/none/absent sides carry none. The visual
 *  default border is a renderer decision the adapter injects — the engine
 *  measures only declared edges. */
function borderEdgePx(edge: LayoutBorderEdge | undefined): number {
  if (edge && edge.style && edge.style !== "nil" && edge.style !== "none" && edge.px != null) {
    return edge.px;
  }
  return 0;
}
