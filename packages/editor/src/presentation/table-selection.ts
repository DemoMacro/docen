import type { Box } from "../drawing/geometry";

/** One projected table cell at its grid slot — the fields the selection model
 *  reads from a projected table member (the rest of the paint payload is
 *  ignored). */
export interface TableCellSlot {
  row: number;
  col: number;
  spanW: number;
  spanH: number;
}

export interface TableSelectionView {
  x: number;
  y: number;
  table: {
    columnWidthsPx: number[];
    rows: { heightPx: number; cells: TableCellSlot[] }[];
  };
}

/** Word's cross-cell selection: anchor and head cells bound a grid block. */
export interface TableSelectionRange {
  anchor: { row: number; col: number };
  head: { row: number; col: number };
}

export interface TableGripHit {
  kind: "row" | "col" | "table";
  index: number;
  /** False for the table-wide hover preview: the square is visible, but only
   *  the corner window clicks it. */
  clickable: boolean;
}

export interface TableCellRect extends Box {
  cell: TableCellSlot;
}

export function tableEdges(member: TableSelectionView): {
  colEdges: number[];
  rowEdges: number[];
} {
  const prefix = (values: number[]): number[] => {
    const edges = [0];
    for (const value of values) edges.push(edges[edges.length - 1]! + Math.max(0, value));
    return edges;
  };
  return {
    colEdges: prefix(member.table.columnWidthsPx),
    rowEdges: prefix(member.table.rows.map((row) => row.heightPx)),
  };
}

function cellRectAt(
  member: TableSelectionView,
  cell: TableCellSlot,
  colEdges = tableEdges(member).colEdges,
  rowEdges = tableEdges(member).rowEdges,
): TableCellRect {
  const x = member.x + colEdges[Math.min(cell.col, colEdges.length - 1)]!;
  const y = member.y + rowEdges[Math.min(cell.row, rowEdges.length - 1)]!;
  const right = member.x + colEdges[Math.min(cell.col + cell.spanW, colEdges.length - 1)]!;
  const bottom = member.y + rowEdges[Math.min(cell.row + cell.spanH, rowEdges.length - 1)]!;
  return { cell, x, y, width: right - x, height: bottom - y };
}

/** The origin cell under a slide-local point; null outside the table. */
export function tableCellAt(
  member: TableSelectionView,
  x: number,
  y: number,
): TableCellRect | null {
  const { colEdges, rowEdges } = tableEdges(member);
  const lx = x - member.x;
  const ly = y - member.y;
  if (
    lx < 0 ||
    ly < 0 ||
    lx >= colEdges[colEdges.length - 1]! ||
    ly >= rowEdges[rowEdges.length - 1]!
  )
    return null;
  const col = colEdges.findLastIndex((edge) => edge <= lx) || 0;
  const row = rowEdges.findLastIndex((edge) => edge <= ly) || 0;
  const cell = member.table.rows
    .flatMap((r) => r.cells)
    .find((c) => c.row <= row && row < c.row + c.spanH && c.col <= col && col < c.col + c.spanW);
  return cell ? cellRectAt(member, cell, colEdges, rowEdges) : null;
}

/** Dragging outside the table clamps to the nearest grid slot, matching a
 *  cross-cell drag that runs past the rim. */
export function tableCellNearAt(
  member: TableSelectionView,
  x: number,
  y: number,
): TableCellRect | null {
  const exact = tableCellAt(member, x, y);
  if (exact) return exact;
  const { colEdges, rowEdges } = tableEdges(member);
  const col = Math.max(
    0,
    Math.min(colEdges.length - 2, colEdges.findLastIndex((edge) => edge <= x - member.x) || 0),
  );
  const row = Math.max(
    0,
    Math.min(rowEdges.length - 2, rowEdges.findLastIndex((edge) => edge <= y - member.y) || 0),
  );
  const cell = member.table.rows
    .flatMap((r) => r.cells)
    .find((c) => c.row <= row && row < c.row + c.spanH && c.col <= col && col < c.col + c.spanW);
  return cell ? cellRectAt(member, cell) : null;
}

function coversSlot(cell: TableCellSlot, row: number, col: number): boolean {
  return (
    cell.row <= row && row < cell.row + cell.spanH && cell.col <= col && col < cell.col + cell.spanW
  );
}

/** Every origin cell crossed by the grid rectangle. A merged origin covers its
 *  whole span, exactly like DOCX's `cellsInRect`. */
export function tableSelectionCells(
  member: TableSelectionView,
  range: TableSelectionRange,
): TableCellSlot[] {
  const cells = member.table.rows.flatMap((row) => row.cells);
  const anchor = cells.find((cell) => coversSlot(cell, range.anchor.row, range.anchor.col));
  const head = cells.find((cell) => coversSlot(cell, range.head.row, range.head.col));
  if (!anchor || !head) return [];
  const colFrom = Math.min(anchor.col, head.col);
  const colTo = Math.max(anchor.col + anchor.spanW, head.col + head.spanW);
  const rowFrom = Math.min(anchor.row, head.row);
  const rowTo = Math.max(anchor.row + anchor.spanH, head.row + head.spanH);
  return cells.filter(
    (cell) =>
      cell.row < rowTo &&
      cell.row + cell.spanH > rowFrom &&
      cell.col < colTo &&
      cell.col + cell.spanW > colFrom,
  );
}

/** One DOCX grip → the cell pair that selects that row/column/table. */
export function tableSelectionFor(
  member: TableSelectionView,
  kind: "row" | "col" | "table",
  index = 0,
): TableSelectionRange | null {
  const columns = member.table.columnWidthsPx.length;
  const rows = member.table.rows.length;
  if (kind === "col")
    return index >= 0 && index < columns
      ? { anchor: { row: 0, col: index }, head: { row: rows - 1, col: index } }
      : null;
  if (kind === "row")
    return index >= 0 && index < rows
      ? { anchor: { row: index, col: 0 }, head: { row: index, col: columns - 1 } }
      : null;
  return rows > 0 && columns > 0
    ? { anchor: { row: 0, col: 0 }, head: { row: rows - 1, col: columns - 1 } }
    : null;
}

/** Cell boxes for painting the selection; document order. */
export function tableSelectionRects(
  member: TableSelectionView,
  range: TableSelectionRange,
): TableCellRect[] {
  return tableSelectionCells(member, range).map((cell) => cellRectAt(member, cell));
}

/** The Word table grip under a member-local point. `hover` keeps the select-all
 *  square visible over the table body; a press requires one of the grip windows
 *  so ordinary cell clicks remain editing gestures. */
export function tableGripAt(
  member: TableSelectionView,
  x: number,
  y: number,
  hover = false,
): TableGripHit | null {
  const { colEdges, rowEdges } = tableEdges(member);
  const right = colEdges[colEdges.length - 1]!;
  const bottom = rowEdges[rowEdges.length - 1]!;
  const lx = x - member.x;
  const ly = y - member.y;
  const inside = lx >= 0 && lx < right && ly >= 0 && ly < bottom;
  if (lx >= -13 && lx <= 1 && ly >= -13 && ly <= 1)
    return { kind: "table", index: 0, clickable: true };
  if (ly >= -14 && ly <= 4 && lx > 0 && lx < right) {
    const index = colEdges.findIndex(
      (edge, index) => index < colEdges.length - 1 && lx >= edge && lx < colEdges[index + 1]!,
    );
    if (index >= 0) return { kind: "col", index, clickable: true };
  }
  if (lx >= -14 && lx <= 4 && ly > 0 && ly < bottom) {
    const index = rowEdges.findIndex(
      (edge, index) => index < rowEdges.length - 1 && ly >= edge && ly < rowEdges[index + 1]!,
    );
    if (index >= 0) return { kind: "row", index, clickable: true };
  }
  return hover && inside ? { kind: "table", index: 0, clickable: false } : null;
}
