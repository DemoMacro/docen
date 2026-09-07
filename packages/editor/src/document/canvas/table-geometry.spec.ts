import { describe, expect, it } from "vitest";

import type { TableZone } from "./caret-map";
import {
  calcColumnResize,
  calcRowResize,
  calcTableScale,
  hitTableElements,
  MIN_COL_TWIPS,
  MIN_ROW_TWIPS,
} from "./table-geometry";

describe("table-geometry", () => {
  const sampleZone: TableZone = {
    page: 0,
    xPx: 100,
    yPx: 150,
    widthPx: 300,
    heightPx: 120,
    colEdges: [0, 100, 200, 300], // 3 columns
    rowEdges: [0, 40, 80, 120], // 3 rows
  };

  describe("hitTableElements", () => {
    it("detects bottom-right corner resize handle", () => {
      // Corner is at (xPx + widthPx, yPx + heightPx) = (400, 270)
      const hit = hitTableElements(sampleZone, 400, 270);
      expect(hit).toEqual({ kind: "corner-handle" });

      const nearHit = hitTableElements(sampleZone, 396, 268);
      expect(nearHit).toEqual({ kind: "corner-handle" });
    });

    it("detects interior column border hit", () => {
      // Border 1 is at x = 100 + 100 = 200
      const hit = hitTableElements(sampleZone, 201, 180);
      expect(hit).toEqual({ kind: "col-border", colIndex: 1 });

      // Border 2 is at x = 100 + 200 = 300
      const hit2 = hitTableElements(sampleZone, 299, 200);
      expect(hit2).toEqual({ kind: "col-border", colIndex: 2 });
    });

    it("detects rightmost column border hit", () => {
      // Border 3 is at x = 100 + 300 = 400, mid-height y = 200
      const hit = hitTableElements(sampleZone, 400, 200);
      expect(hit).toEqual({ kind: "col-border", colIndex: 3 });
    });

    it("detects interior row border hit", () => {
      // Row border 1 is at y = 150 + 40 = 190
      const hit = hitTableElements(sampleZone, 150, 191);
      expect(hit).toEqual({ kind: "row-border", rowIndex: 1 });

      // Row border 2 is at y = 150 + 80 = 230
      const hit2 = hitTableElements(sampleZone, 250, 230);
      expect(hit2).toEqual({ kind: "row-border", rowIndex: 2 });
    });

    it("detects quick-insert column button zone", () => {
      // Border 1 top is at x = 200, y = 150. Above top edge: y = 140
      const hit = hitTableElements(sampleZone, 200, 140);
      expect(hit).toEqual({ kind: "quick-col", colIndex: 1 });
    });

    it("detects quick-insert row button zone", () => {
      // Border 1 left is at x = 100, y = 190. Left of left edge: x = 90
      const hit = hitTableElements(sampleZone, 90, 190);
      expect(hit).toEqual({ kind: "quick-row", rowIndex: 1 });
    });

    it("returns null for points in the middle of cells", () => {
      const hit = hitTableElements(sampleZone, 150, 170);
      expect(hit).toBeNull();
    });
  });

  describe("calcColumnResize", () => {
    it("redistributes width between two interior columns without changing total width", () => {
      const initial = [2000, 2000, 2000];
      // Drag border 1 (between col 0 and col 1) rightward by +500 twips
      const next = calcColumnResize(initial, 1, 500);
      expect(next).toEqual([2500, 1500, 2000]);
      expect(next.reduce((a, b) => a + b, 0)).toBe(6000);
    });

    it("clamps column width so neither column drops below MIN_COL_TWIPS", () => {
      const initial = [2000, 2000, 2000];
      // Try to drag border 1 rightward by +3000 twips (which would make col 1 negative)
      const next = calcColumnResize(initial, 1, 3000);
      expect(next[1]).toBe(MIN_COL_TWIPS);
      expect(next[0]).toBe(4000 - MIN_COL_TWIPS);
      expect(next[2]).toBe(2000);
    });

    it("resizes last column when dragging outer right border", () => {
      const initial = [2000, 2000, 2000];
      // Drag border 3 (outer right) rightward by +800 twips
      const next = calcColumnResize(initial, 3, 800);
      expect(next).toEqual([2000, 2000, 2800]);
    });

    it("shifts subsequent columns when shiftKey is held", () => {
      const initial = [2000, 2000, 2000];
      // Drag border 1 rightward by +400 twips with Shift
      const next = calcColumnResize(initial, 1, 400, { shiftKey: true });
      expect(next).toEqual([2400, 2000, 2000]);
    });
  });

  describe("calcRowResize", () => {
    it("returns atLeast rule by default with updated value", () => {
      const res = calcRowResize(400, 200);
      expect(res).toEqual({ rule: "atLeast", value: 600 });
    });

    it("returns exact rule when isAlt is true", () => {
      const res = calcRowResize(400, 150, { isAlt: true });
      expect(res).toEqual({ rule: "exact", value: 550 });
    });

    it("clamps minimum row height", () => {
      const res = calcRowResize(400, -800);
      expect(res.value).toBe(MIN_ROW_TWIPS);
    });
  });

  describe("calcTableScale", () => {
    it("scales all columns proportionally", () => {
      const initial = [1000, 2000, 3000];
      const scaled = calcTableScale(initial, 1.5);
      expect(scaled).toEqual([1500, 3000, 4500]);
    });

    it("clamps each column to minWidth", () => {
      const initial = [1000, 2000];
      const scaled = calcTableScale(initial, 0.1);
      expect(scaled).toEqual([MIN_COL_TWIPS, MIN_COL_TWIPS]);
    });
  });
});
