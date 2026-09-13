// @vitest-environment node
import type { GeometryGuide } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { presetShapePaths, presetShapeTextRect } from "./preset-shape";
import { PRESET_SHAPE_DEFS } from "./preset-shape-data";

// Golden d strings are hand-derived from the ECMA-376 definitions (100×100
// boxes, spec defaults) — an independent reference for the evaluator, its
// built-in guide set, and the arcTo→cubic conversion.
describe("presetShapePaths", () => {
  it("evaluates polygon presets through the guide table", () => {
    expect(presetShapePaths("triangle", 100, 100)).toEqual([
      { d: "M 0 100 L 50 0 L 100 100 Z", fill: true, stroke: true },
    ]);
    expect(presetShapePaths("rect", 100, 50)?.[0].d).toBe("M 0 0 L 100 0 L 100 50 L 0 50 Z");
    expect(presetShapePaths("line", 100, 100)?.[0].d).toBe("M 0 0 L 100 100");
    expect(presetShapePaths("rtTriangle", 100, 100)?.[0].d).toBe("M 0 100 L 0 0 L 100 100 Z");
  });

  it("converts arcTo to cubic segments from the center the current point implies", () => {
    // roundRect 16667: k = 4/3·tan(22.5°) · 16.667 = 9.20 → corner controls at 7.46/92.54.
    expect(presetShapePaths("roundRect", 100, 100)?.[0].d).toBe(
      "M 0 16.67 C 0 7.46 7.46 0 16.67 0 L 83.33 0 C 92.54 0 100 7.46 100 16.67" +
        " L 100 83.33 C 100 92.54 92.54 100 83.33 100 L 16.67 100 C 7.46 100 0 92.54 0 83.33 Z",
    );
    // pie 0°→270°: arc from (100,50), then a line to the center, closed.
    expect(presetShapePaths("pie", 100, 100)?.[0].d).toBe(
      "M 100 50 C 100 77.61 77.61 100 50 100 C 22.39 100 0 77.61 0 50 C 0 22.39 22.39 0 50 0 L 50 50 Z",
    );
  });

  it("keeps negative sweeps and subpath winding (donut ring)", () => {
    expect(presetShapePaths("donut", 100, 100)?.[0].d).toBe(
      "M 0 50 C 0 22.39 22.39 0 50 0 C 77.61 0 100 22.39 100 50 C 100 77.61 77.61 100 50 100" +
        " C 22.39 100 0 77.61 0 50 Z" +
        " M 25 50 C 25 63.81 36.19 75 50 75 C 63.81 75 75 63.81 75 50 C 75 36.19 63.81 25 50 25" +
        " C 36.19 25 25 36.19 25 50 Z",
    );
  });

  it("applies adjustmentValues over the preset defaults by name", () => {
    const wedge: GeometryGuide[] = [
      { name: "adj1", formula: "val 5400000" },
      { name: "adj2", formula: "val 10800000" },
    ];
    expect(presetShapePaths("pie", 100, 100, wedge)?.[0].d).toBe(
      "M 50 100 C 22.39 100 0 77.61 0 50 L 50 50 Z",
    );
    expect(
      presetShapePaths("triangle", 100, 100, [{ name: "adj", formula: "val 25000" }])?.[0].d,
    ).toBe("M 0 100 L 25 0 L 100 100 Z");
  });

  it("splits multi-path shapes into fill/stroke outlines and drops shading layers", () => {
    // can: silhouette (fill only), a lighten shading layer (dropped), and the
    // rim (stroke only, fill "none").
    const outlines = presetShapePaths("can", 100, 100);
    expect(outlines).toHaveLength(2);
    expect(outlines?.map((o) => [o.fill, o.stroke])).toEqual([
      [true, false],
      [false, true],
    ]);
  });

  it("scales path coordinate spaces to the box (cloudCallout's 43200 space)", () => {
    const d = presetShapePaths("cloudCallout", 100, 100)?.[0].d;
    expect(d).toBeDefined();
    const coords = d!.match(/-?\d+(\.\d+)?/g)!.map(Number);
    // The balloon outline may poke slightly past the box (Word renders the
    // same excursion), but an unscaled path would carry raw 43200-space
    // literals — a factor-432 excursion.
    expect(Math.max(...coords)).toBeLessThanOrEqual(110);
    expect(Math.min(...coords)).toBeGreaterThanOrEqual(-10);
    // Unscaled output would carry raw 43200-space literals — a factor-432 excursion.
  });

  it("tolerates double spaces in the ECMA guide formulas (star5's svc)", () => {
    // The spec source separates two operands with two spaces; a naive split
    // would coerce the empty token to 0 and collapse the star's hub.
    expect(presetShapePaths("star5", 100, 100)?.[0].d).toMatch(/^M 0 38\.2 L 38\.2 38\.2 L 50 0 /);
  });

  it("returns undefined for unknown presets", () => {
    expect(presetShapePaths("bogus", 100, 100)).toBeUndefined();
    expect(presetShapePaths("textNoShape", 100, 100)).toBeUndefined();
  });

  it("evaluates every defined preset to finite, well-formed path data", () => {
    const presets = Object.keys(PRESET_SHAPE_DEFS);
    expect(presets.length).toBeGreaterThanOrEqual(186);
    for (const preset of presets) {
      const outlines = presetShapePaths(preset, 100, 100);
      expect(outlines, preset).toBeDefined();
      for (const { d } of outlines!) {
        expect(d, preset).toMatch(/^[MLCQZ][MLCQZ\s\d.-]*$/);
        expect(d, preset).not.toMatch(/NaN|Infinity/);
      }
      // Degenerate boxes must not throw or produce non-finite coordinates.
      expect(presetShapePaths(preset, 0, 0), preset).toBeDefined();
    }
  });
});

describe("presetShapeTextRect", () => {
  it("evaluates the ellipse's inscribed text rectangle", () => {
    // The 45-degree contact points: each side insets w*(1-cos45°)/2 = 14.64
    // at 100px — Word keeps an ellipse's words off the rim.
    const tr = presetShapeTextRect("ellipse", 100, 100)!;
    expect(tr.l).toBeCloseTo((100 * (1 - Math.SQRT2 / 2)) / 2, 9);
    expect(tr.t).toBeCloseTo(tr.l, 9);
    expect(tr.r).toBeCloseTo(100 - tr.l, 9);
    expect(tr.b).toBeCloseTo(100 - tr.t, 9);
  });

  it("returns undefined for presets without a rect definition", () => {
    expect(presetShapeTextRect("rect", 100, 50)).toBeUndefined();
    expect(presetShapeTextRect("bogus", 100, 50)).toBeUndefined();
  });
});
