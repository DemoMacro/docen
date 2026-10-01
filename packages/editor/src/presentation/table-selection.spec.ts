import { describe, expect, it } from "vitest";

import {
  tableCellAt,
  tableCellNearAt,
  tableSelectionRects,
  tableSelectionCells,
  tableSelectionFor,
  type TableSelectionView,
} from "./table-selection";

const member: TableSelectionView = {
  x: 20,
  y: 30,
  table: {
    columnWidthsPx: [10, 20, 30],
    rows: [
      {
        heightPx: 10,
        cells: [
          { row: 0, col: 0, spanW: 1, spanH: 1 },
          { row: 0, col: 1, spanW: 2, spanH: 1 },
        ],
      },
      {
        heightPx: 20,
        cells: [
          { row: 1, col: 0, spanW: 2, spanH: 1 },
          { row: 1, col: 2, spanW: 1, spanH: 1 },
        ],
      },
    ],
  },
};

describe("pptx table selection", () => {
  it("maps grid slots to origin cells and painted boxes", () => {
    expect(tableCellAt(member, 20, 30)?.cell).toEqual({ row: 0, col: 0, spanW: 1, spanH: 1 });
    expect(tableCellAt(member, 41, 31)?.cell).toEqual({ row: 0, col: 1, spanW: 2, spanH: 1 });
    expect(tableCellAt(member, 21, 51)?.cell).toEqual({ row: 1, col: 0, spanW: 2, spanH: 1 });
    expect(tableCellAt(member, 10, 40)).toBeNull();
    expect(
      tableSelectionRects(member, {
        anchor: { row: 0, col: 0 },
        head: { row: 0, col: 0 },
      }),
    ).toEqual([
      {
        cell: { row: 0, col: 0, spanW: 1, spanH: 1 },
        x: 20,
        y: 30,
        width: 10,
        height: 10,
      },
    ]);
  });

  it("extends a cross-cell drag across whole merged slots", () => {
    const cells = tableSelectionCells(member, {
      anchor: { row: 0, col: 0 },
      head: tableCellNearAt(member, 100, 100)!.cell,
    });
    expect(cells).toEqual([
      { row: 0, col: 0, spanW: 1, spanH: 1 },
      { row: 0, col: 1, spanW: 2, spanH: 1 },
      { row: 1, col: 0, spanW: 2, spanH: 1 },
      { row: 1, col: 2, spanW: 1, spanH: 1 },
    ]);
  });

  it("creates Word-style row, column, and table grips", () => {
    expect(tableSelectionCells(member, tableSelectionFor(member, "col", 2)!)).toEqual([
      { row: 0, col: 1, spanW: 2, spanH: 1 },
      { row: 1, col: 0, spanW: 2, spanH: 1 },
      { row: 1, col: 2, spanW: 1, spanH: 1 },
    ]);
    expect(tableSelectionCells(member, tableSelectionFor(member, "row", 1)!)).toEqual([
      { row: 1, col: 0, spanW: 2, spanH: 1 },
      { row: 1, col: 2, spanW: 1, spanH: 1 },
    ]);
    expect(tableSelectionCells(member, tableSelectionFor(member, "table")!)).toHaveLength(4);
  });
});
