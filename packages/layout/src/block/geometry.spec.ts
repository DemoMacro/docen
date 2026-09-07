// The shared geometry both the painter and the caret map consume — each
// function is the single authority for its sum, so the tests pin the sums
// (not a rendering outcome).

import { describe, expect, it } from "vitest";

import type { LaidOutLine, LaidOutParagraph, LaidOutTable } from "../layout-result";
import {
  gridPadOf,
  justifiedIntervals,
  lineBaselineDepthPx,
  lineOriginXPx,
  tableGridOf,
} from "./geometry";

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

describe("lineBaselineDepthPx", () => {
  it("sinks the baseline to the picture bottom on a picture-floored line", () => {
    // Word treats an inline picture as a single big character whose bottom
    // edge sits ON the baseline: the floored line's baseline = pad + picture.
    // Grid span 28 centers the 21.5px picture (pad 3.25) → 3.25 + 21.5.
    expect(
      lineBaselineDepthPx(
        line({ grid: true, pictureFloored: true, heightPx: 28, naturalPx: 21.5, textEmPx: 7 }),
      ),
    ).toBeCloseTo(24.75, 5);
    // A text-only line keeps the text ascent depth — the grid pad centers the
    // text em here (pad (28−7)/2 = 10.5, plus 0.85 × 7).
    expect(
      lineBaselineDepthPx(line({ grid: true, heightPx: 28, naturalPx: 21.5, textEmPx: 7 })),
    ).toBeCloseTo(16.45, 5);
    // A non-grid floored line is top-anchored, so the baseline is the picture.
    expect(
      lineBaselineDepthPx(line({ pictureFloored: true, heightPx: 28, naturalPx: 21.5 })),
    ).toBeCloseTo(21.5, 5);
  });
});

describe("gridPadOf", () => {
  it("centers the natural box in a grid span and pins everything else to the top", () => {
    expect(gridPadOf(line({ grid: true, heightPx: 34, naturalPx: 18 }))).toBe(8);
    expect(gridPadOf(line({ grid: true, heightPx: 10, naturalPx: 18 }))).toBe(0);
    // Textless non-grid lines (strut rows) have no ascent to scale.
    expect(gridPadOf(line({ heightPx: 34, naturalPx: 18 }))).toBe(0);
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

  it("splits a non-grid multiple line's slack on the natural box's ascent ratio", () => {
    // Word's multiple spacing scales the whole box: the baseline rides at
    // factor × its single-line height (0.85 × em here), so of the 8px of
    // slack the ascent share (8.5/20) lands above the glyphs — not all 8px
    // below them.
    expect(gridPadOf(line({ heightPx: 28, naturalPx: 20, textEmPx: 10 }))).toBeCloseTo(3.4, 5);
    // A picture-floored line keeps its box top-anchored (the picture sits on
    // the baseline, i.e. the box bottom).
    expect(
      gridPadOf(line({ heightPx: 28, naturalPx: 20, textEmPx: 10, pictureFloored: true })),
    ).toBe(0);
    // Single spacing has no slack to split.
    expect(gridPadOf(line({ heightPx: 20, naturalPx: 20, textEmPx: 10 }))).toBe(0);
  });

  it("keeps atLeast top-anchored and bottoms an exact line (Word's regimes)", () => {
    expect(
      gridPadOf(line({ heightPx: 28, naturalPx: 20, textEmPx: 10, spacingRule: "atLeast" })),
    ).toBe(0);
    // exact sinks the glyphs onto the box bottom — the slack rides above.
    expect(
      gridPadOf(line({ heightPx: 28, naturalPx: 20, textEmPx: 10, spacingRule: "exact" })),
    ).toBe(8);
    // Undersized exact keeps the natural position (pad clamps at 0): Word
    // clips the glyph tops at the box edge rather than lifting the text
    // (pixel-verified against the reference render).
    expect(
      gridPadOf(line({ heightPx: 16, naturalPx: 20, textEmPx: 10, spacingRule: "exact" })),
    ).toBe(0);
  });

  it("half-leads a textbox grid line the same as a body line — compatLnSpc changes nothing", () => {
    expect(gridPadOf(line({ grid: true, heightPx: 45.3, textEmPx: 18.7 }))).toBeCloseTo(13.3, 5);
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
