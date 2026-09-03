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
  LayoutBlock,
  LayoutBlockContext,
  LayoutBorderEdge,
  LayoutCellInsets,
  LayoutTable,
} from "../layout-doc";
import type { LaidOutCell, LaidOutTable } from "../layout-result";
import { minWidthOfInline } from "../text/line-break";
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
  // w:tblLayout autofit re-fits the columns to their content (and, for an
  // auto-width table, the table to its columns) at every layout; the fixed
  // path just scales the grid. Absent layout = fixed (see LayoutTable).
  const { columnWidths, tableWidth } =
    table.layout === "autofit"
      ? autofitColumns(table, containerWidth, measurer)
      : { columnWidths: tableColumnWidths(table, widthPx), tableWidth: widthPx };

  // Word honors w:tblHeader only as a contiguous prefix from the first row —
  // marks past the prefix are voided here so no later page can re-derive a
  // band from a marked row that is no longer at a table top.
  let bandEnd = 0;
  while (bandEnd < table.rows.length && table.rows[bandEnd].tableHeader) bandEnd++;

  let heightPx = 0;
  // Columns a row-spanning cell above already owns — vertical merges drop
  // their continuation cells at projection, so each row's walk must skip the
  // columns a span anchored on an earlier row keeps, or the row's cells pack
  // one column left (wrong widths and grid-rim fallbacks follow).
  const occupied: boolean[][] = table.rows.map(() => []);
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
    // w:trHeight atLeast floors the row's CONTENT box (see the cell loop);
    // exact fixes the row outright.
    const tr = row.height;
    const trFloorPx = tr && tr.rule === "atLeast" ? tr.px : 0;
    // (natural content + vertical overhead per cell, for the vAlign slack
    // pass after the row height settles)
    const totals: { stackedPx: number; overheadPx: number }[] = [];
    const cells: LaidOutCell[] = row.cells.map((cell) => {
      const colspan = cell.colspan ?? 1;
      while (colCursor < columnWidths.length && occupied[rowIndex]![colCursor]) colCursor++;
      let cellWidth = 0;
      for (let i = 0; i < colspan && colCursor + i < columnWidths.length; i++) {
        cellWidth += columnWidths[colCursor + i];
      }
      for (let dr = 0; dr < (cell.rowspan ?? 1) && rowIndex + dr < table.rows.length; dr++)
        for (let dc = 0; dc < colspan && colCursor + dc < columnWidths.length; dc++)
          occupied[rowIndex + dr]![colCursor + dc] = true;
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
      const { left: leftEdge, right: rightEdge } = horizontalEdgesOf(
        cell,
        table,
        firstCol,
        lastCol,
      );
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

      // Word applies trHeight to the CONTENT box alone (verified: a 374-twip
      // atLeast row with 85-twip top/bottom cell margins renders 18.7pt +
      // 8.5pt + border — the margins ride on TOP of the trHeight match, they
      // do not join the content in competing against it), then adds the
      // margins and border to whatever won.
      const contentPx = Math.max(stacked.heightPx, trFloorPx);
      const cellHeightPx = contentPx + vOverheadPx;
      totals.push({ stackedPx: stacked.heightPx, overheadPx: vOverheadPx });
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

    const trExactPx = tr && tr.rule === "exact" ? tr.px : undefined;
    if (trExactPx != null) rowHeight = trExactPx;
    // w:vAlign: place the content in the row's slack (read back from the
    // source cells — the laid-out cell carries only the resolved offset). The
    // slack is measured against the cell's NATURAL content (stack + overhead):
    // an atLeast trHeight floor makes the floored height the row's tallest,
    // and centering against it would report zero slack in a row with real
    // room (corpus: the honor table's header row). A vertically merged cell
    // keeps its content on the start row (the span-distribution boundary
    // above), so only single-row cells shift.
    cells.forEach((cell, i) => {
      const nat = totals[i];
      const slack = nat ? rowHeight - nat.stackedPx - nat.overheadPx : 0;
      const va = row.cells[i]?.verticalAlign;
      if (slack <= 0 || (cell.rowspan ?? 1) > 1) return;
      cell.contentOffsetYPx = va === "center" ? slack / 2 : va === "bottom" ? slack : undefined;
    });
    heightPx += rowHeight;
    return {
      heightPx: rowHeight,
      cells,
      tableHeader: rowIndex < bandEnd || undefined,
      cantSplit: row.cantSplit || undefined,
      exactHeight: tr?.rule === "exact" || undefined,
    };
  });

  return {
    kind: "table",
    widthPx: tableWidth,
    columnWidthsPx: columnWidths,
    // w:jc: the table box's placement in the flow column — center/right
    // against the column width, which goes negative for a table wider than
    // the column (Word centers those into the margins).
    offsetXPx:
      table.align === "center"
        ? (containerWidth - tableWidth) / 2
        : table.align === "right"
          ? containerWidth - tableWidth
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
  const grid = rawGridOf(table);
  if (grid.length === 0) return [];
  const total = grid.reduce((a, b) => a + b, 0) || 1;
  return grid.map((w) => (w / total) * tableWidth);
}

/** The raw tblGrid (or the first row's cell widths), unscaled, px. */
function rawGridOf(table: LayoutTable): number[] {
  if (table.columnWidthsPx && table.columnWidthsPx.length > 0) {
    return [...table.columnWidthsPx];
  }
  const grid: number[] = [];
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
  return grid;
}

/** Word's autofit (w:tblLayout autofit): every column starts at its grid
 *  preference and grows only when a cell cannot wrap inside it — the fit
 *  measures each cell wrapped AT its grid width, so prose keeps the column
 *  and only an unbreakable run (a long word, an image) widens it past the
 *  grid. A pinned w:tblW scales the result; an auto table sizes to its
 *  columns, never beyond what the grid already spans. Spanning cells don't
 *  constrain the fit (Word distributes across the span; a per-column split
 *  would be guesswork). */
function autofitColumns(
  table: LayoutTable,
  containerWidth: number,
  measurer: TextMeasurer,
): { columnWidths: number[]; tableWidth: number } {
  const grid = rawGridOf(table);
  if (grid.length === 0) return { columnWidths: [], tableWidth: containerWidth };
  const content = grid.map(() => 0);
  // Same occupancy walk the fixed path uses: merged rows drop their
  // continuation cells, so each row's cells must skip columns a span above
  // keeps or their content floor lands on the wrong column.
  const occupied: boolean[][] = table.rows.map(() => []);
  for (const [rowIndex, row] of table.rows.entries()) {
    let col = 0;
    for (const cell of row.cells) {
      const span = cell.colspan ?? 1;
      while (col < grid.length && occupied[rowIndex]![col]) col++;
      if (span === 1 && col < content.length) {
        const insets = mergeInsets(cell.insets, table.cellInsets);
        // The same table-level border fallback the wrap width uses below —
        // a table-bordered cell (Table Grid) must charge its side borders
        // to the column or the fit comes out a hair too narrow.
        const { left, right } = horizontalEdgesOf(
          cell,
          table,
          col === 0,
          col + span >= grid.length,
        );
        const frame = (insets.left ?? 0) + (insets.right ?? 0) + edgeWidth(left) + edgeWidth(right);
        // A column without a grid entry measures unconstrained — content
        // starts from its own unwrapped width.
        const budget = grid[col]! > 0 ? grid[col]! - frame : 1e9;
        const floor = minWidthOfBlocks(cell.blocks, measurer, budget) + frame;
        content[col] = Math.max(content[col]!, floor);
      }
      for (let dr = 0; dr < (cell.rowspan ?? 1) && rowIndex + dr < table.rows.length; dr++)
        for (let dc = 0; dc < span && col + dc < grid.length; dc++)
          occupied[rowIndex + dr]![col + dc] = true;
      col += span;
    }
  }
  // The grid preference is the starting point — content only widens a
  // column (a 0/absent grid entry starts from the content itself).
  const target = grid.map((w, c) => Math.max(w > 0 ? w : content[c]!, content[c]!));
  const total = target.reduce((a, b) => a + b, 0);
  if (total <= 0) return { columnWidths: grid, tableWidth: containerWidth };
  const width =
    table.width?.type === "percent"
      ? (containerWidth * table.width.percent) / 100
      : table.width?.type === "px"
        ? table.width.px
        : total;
  return { columnWidths: target.map((w) => (w / total) * width), tableWidth: width };
}

/** A cell stack's widest line when wrapped at `budget` (hard breaks split it;
 *  nested tables contribute their grid sum; placeholders nothing). */
function minWidthOfBlocks(
  blocks: readonly LayoutBlock[],
  measurer: TextMeasurer,
  budget: number,
): number {
  let widest = 0;
  for (const block of blocks) {
    switch (block.kind) {
      case "paragraph":
        widest = Math.max(widest, minWidthOfInline(block.inline, measurer, budget));
        break;
      case "group":
        widest = Math.max(widest, minWidthOfBlocks(block.blocks, measurer, budget));
        break;
      case "table": {
        const sum = (block.columnWidthsPx ?? []).reduce((a, b) => a + b, 0);
        widest = Math.max(widest, sum);
        break;
      }
      default:
        break;
    }
  }
  return widest;
}

/** The border edge's painted width (nil/none → 0). */
function edgeWidth(edge: LayoutBorderEdge | undefined): number {
  if (!edge || edge.style === "nil" || edge.style === "none") return 0;
  return edge.px ?? 0;
}

/** A cell's horizontal border edges after the table-level fallback — the
 *  grid rim (first/last column) takes the outer edges, interior boundaries
 *  the inside one (CT_TblBorders semantics). Single source for the wrap
 *  width and the autofit frame so both charge the same borders. */
function horizontalEdgesOf(
  cell: LayoutTable["rows"][number]["cells"][number],
  table: LayoutTable,
  firstCol: boolean,
  lastCol: boolean,
): { left: LayoutBorderEdge | undefined; right: LayoutBorderEdge | undefined } {
  return {
    left: cell.borders?.left ?? (firstCol ? table.borders?.left : table.borders?.insideVertical),
    right: cell.borders?.right ?? (lastCol ? table.borders?.right : table.borders?.insideVertical),
  };
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
