// The shared geometry both the painter and the caret map consume — each
// function is the single authority for its sum, so the tests pin the sums
// (not a rendering outcome).

import { describe, expect, it } from "vitest";

import type { LaidOutLine, LaidOutParagraph, LaidOutTable } from "../layout-result";
import { gridPadOf, justifiedIntervals, lineOriginXPx, tableGridOf } from "./geometry";

const line = (over: Partial<LaidOutLine> = {}): LaidOutLine =>
  ({
    yPx: 0,
    heightPx: 20,
    naturalPx: 16,
    endInlineIndex: 0,
    items: [],
    ...over,
  }) as LaidOutLine;

const para = (over: Partial<LaidOutParagraph> = {}): LaidOutParagraph =>
  ({
    kind: "paragraph",
    heightPx: 20,
    beforePx: 0,
    afterPx: 0,
    lines: [],
    inline: [],
    ...over,
  }) as LaidOutParagraph;

describe("lineOriginXPx", () => {
  it("sums the left indent, the line's own first-line flag, and the float shift", () => {
    const p = para({ indent: { leftPx: 24, firstLinePx: 21 } });
    expect(lineOriginXPx(p, line({ firstLineIndentPx: 21, xOffsetPx: 30 }))).toBe(75);
    // A continuation line carries no first-line indent.
    expect(lineOriginXPx(p, line({ xOffsetPx: 30 }))).toBe(54);
  });
});

describe("gridPadOf", () => {
  it("centers the natural box in a grid span and pins everything else to the top", () => {
    expect(gridPadOf(line({ grid: true, heightPx: 34, naturalPx: 18 }))).toBe(8);
    expect(gridPadOf(line({ heightPx: 34, naturalPx: 18 }))).toBe(0);
    expect(gridPadOf(line({ grid: true, heightPx: 10, naturalPx: 18 }))).toBe(0);
  });

  it("centers a picture-floored grid line on the picture box, not the text em", () => {
    expect(
      gridPadOf(
        line({ grid: true, pictureFloored: true, heightPx: 28, naturalPx: 21.5, textEmPx: 7 }),
      ),
    ).toBe(3.25);
    // Without the floor the text em keeps winning.
    expect(gridPadOf(line({ grid: true, heightPx: 28, naturalPx: 21.5, textEmPx: 7 }))).toBe(10.5);
  });

  it("half-leads a textbox grid line the same as a body line — compatLnSpc changes nothing", () => {
    expect(gridPadOf(line({ grid: true, heightPx: 45.3, textEmPx: 18.7 }))).toBeCloseTo(13.3, 5);
    // No grid — the slack sinks below the glyphs.
    expect(gridPadOf(line({ heightPx: 45.3, textEmPx: 18.7 }))).toBe(0);
  });
});

describe("justifiedIntervals", () => {
  it("stretches the last item past the wrap width by the hang, earlier ones to the next item's x", () => {
    const a = { kind: "text", inlineIndex: 0, text: "aa", xPx: 0, widthPx: 20 } as const;
    const b = { kind: "text", inlineIndex: 0, text: "bb", xPx: 50, widthPx: 20 } as const;
    const c = { kind: "text", inlineIndex: 0, text: "cc", xPx: 90, widthPx: 20 } as const;
    expect(
      justifiedIntervals(line({ items: [a, b, c], maxWidthPx: 100, justifyGapPx: 2, hangPx: 6 })),
    ).toEqual([50, 90, 106]);
    expect(justifiedIntervals(line({ items: [a] }))).toBeNull();
  });
});

describe("tableGridOf", () => {
  it("walks spans into an occupancy grid and anchors content at insets plus the vAlign offset", () => {
    const cell = (over: Record<string, unknown>): any => ({
      colspan: 1,
      rowspan: 1,
      insets: {},
      stack: [],
      ...over,
    });
    const merged = cell({
      colspan: 2,
      rowspan: 2,
      insets: { left: 8, top: 4 },
      contentOffsetYPx: 12,
      stack: [],
    });
    const table = {
      kind: "table",
      widthPx: 300,
      columnWidthsPx: [100, 100, 100],
      heightPx: 120,
      rows: [
        { heightPx: 60, cells: [merged] },
        { heightPx: 60, cells: [] },
        { heightPx: 40, cells: [cell({ stack: [] }), cell({ stack: [] }), cell({ stack: [] })] },
      ],
    } as unknown as LaidOutTable;
    const grid = tableGridOf(table);
    expect(grid.colX).toEqual([0, 100, 200, 300]);
    expect(grid.rowY).toEqual([0, 60, 120, 160]);
    // The merged cell occupies slots (0,0)-(1,1); row 1's walk skips past it.
    expect(grid.cells[0]).toMatchObject({ col: 0, row: 0, spanW: 2, spanH: 2 });
    expect(grid.occ[1]![0]).toBe(merged);
    expect(grid.occ[1]![2]).toBeUndefined();
    // Row 2's cells start at column 0 — the occupancy walk advanced there.
    expect(grid.cells.slice(1).map((p) => p.col)).toEqual([0, 1, 2]);
    expect(grid.cells[0]).toMatchObject({ contentXPx: 8, contentYPx: 16 });
  });
});
