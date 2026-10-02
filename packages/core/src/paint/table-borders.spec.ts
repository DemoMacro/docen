// @vitest-environment node
import type { LayoutBorderEdge } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { collapsedBorderSegments } from "./table-borders";

const edge = (px: number, extra: Partial<LayoutBorderEdge> = {}): LayoutBorderEdge => ({
  px,
  ...extra,
});
const cell = (
  row: number,
  col: number,
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>,
) => ({ row, col, spanW: 1, spanH: 1, ...(borders ? { borders } : {}) });
const grid = (rows: number, cols: number) => ({
  colX: Array.from({ length: cols + 1 }, (_, i) => i * 100),
  rowY: Array.from({ length: rows + 1 }, (_, i) => i * 40),
});

describe("collapsedBorderSegments", () => {
  it("resolves style-less cell borders — the pptx table-style shape", () => {
    const { colX, rowY } = grid(2, 2);
    const cells = [
      cell(0, 0, { top: edge(1), right: edge(1), bottom: edge(1), left: edge(1) }),
      cell(0, 1, { top: edge(1), right: edge(1), bottom: edge(1), left: edge(1) }),
      cell(1, 0, { top: edge(1), right: edge(1), bottom: edge(1), left: edge(1) }),
      cell(1, 1, { top: edge(1), right: edge(1), bottom: edge(1), left: edge(1) }),
    ];
    const segments = collapsedBorderSegments(colX, rowY, cells);
    expect(segments.filter((s) => s.horizontal)).toHaveLength(3);
    expect(segments.filter((s) => !s.horizontal)).toHaveLength(3);
    expect(segments.every((s) => s.edge.px === 1)).toBe(true);
  });

  it("keeps an explicit nil cell edge silent without falling back to the grid default", () => {
    const { colX, rowY } = grid(2, 1);
    const cells = [cell(0, 0, { bottom: { px: 1, style: "nil" } }), cell(1, 0)];
    const segments = collapsedBorderSegments(colX, rowY, cells, { insideHorizontal: edge(2) });
    expect(segments.filter((s) => s.horizontal && s.y === rowY[1])).toHaveLength(0);
  });

  it("falls back to table-level defaults when cells stay silent", () => {
    const { colX, rowY } = grid(2, 2);
    const segments = collapsedBorderSegments(
      colX,
      rowY,
      [cell(0, 0), cell(0, 1), cell(1, 0), cell(1, 1)],
      { top: edge(3), insideHorizontal: edge(1) },
    );
    expect(segments.filter((s) => s.horizontal && s.y === rowY[0])).toEqual([
      { x: 0, y: 0, lengthPx: 200, horizontal: true, edge: edge(3) },
    ]);
    expect(segments.filter((s) => s.horizontal && s.y === rowY[1])).toEqual([
      { x: 0, y: 40, lengthPx: 200, horizontal: true, edge: edge(1) },
    ]);
  });

  it("draws nothing inside a merged span", () => {
    const { colX, rowY } = grid(2, 1);
    const merged = {
      row: 0,
      col: 0,
      spanW: 1,
      spanH: 2,
      borders: { top: edge(1), bottom: edge(1), left: edge(1), right: edge(1) },
    };
    const segments = collapsedBorderSegments(colX, rowY, [merged]);
    expect(segments.filter((s) => s.horizontal && s.y === rowY[1])).toHaveLength(0);
  });
});
