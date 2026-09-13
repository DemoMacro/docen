import { EMU_PER_PX } from "@docen/layout";
import type { SlideOptions } from "@docen/pptx";
// @vitest-environment node
import { describe, expect, it } from "vitest";

import { hitSlide, offsetChild, resizeChild, slideHits } from "./hit-test";

const px = (v: number): number => Math.round(v * EMU_PER_PX);

describe("slideHits", () => {
  it("builds boxes for every paintable variant", () => {
    const slide: SlideOptions = {
      children: [
        { shape: { x: px(10), y: px(20), width: px(100), height: px(50) } },
        {
          line: { x1: px(0), y1: px(30), x2: px(40), y2: px(10) },
        },
        { picture: { x: px(5), y: px(5), width: px(8), height: px(8), data: "", type: "png" } },
        // Unpainted variants contribute no hit box.
        { table: { rows: [] } },
      ],
    } as unknown as SlideOptions;
    const hits = slideHits(slide);
    expect(hits.map((h) => h.child)).toEqual([0, 1, 2]);
    expect(hits[0]!.box).toEqual({ x: 10, y: 20, width: 100, height: 50 });
    // A line's box spans its endpoints' bounding rectangle.
    expect(hits[1]!.box).toEqual({ x: 0, y: 10, width: 40, height: 20 });
  });
});

describe("hitSlide", () => {
  const slide: SlideOptions = {
    children: [
      { shape: { x: px(0), y: px(0), width: px(100), height: px(100) } },
      { shape: { x: px(50), y: px(50), width: px(100), height: px(100) } },
    ],
  } as SlideOptions;

  it("returns the topmost (latest) child under the point", () => {
    expect(hitSlide(slideHits(slide), 75, 75)).toBe(1);
    expect(hitSlide(slideHits(slide), 25, 25)).toBe(0);
  });

  it("misses outside every box", () => {
    expect(hitSlide(slideHits(slide), 200, 200)).toBe(-1);
  });
});

describe("offsetChild", () => {
  it("shifts transform children in EMU", () => {
    const child = { shape: { x: px(10), y: px(20), width: px(30), height: px(40) } };
    offsetChild(child, 5, -8);
    expect(child.shape.x).toBe(px(15));
    expect(child.shape.y).toBe(px(12));
    expect(child.shape.width).toBe(px(30));
  });

  it("shifts line endpoints together", () => {
    const child = { line: { x1: px(0), y1: px(10), x2: px(50), y2: px(30) } };
    offsetChild(child, 100, 0);
    expect(child.line.x1).toBe(px(100));
    expect(child.line.x2).toBe(px(150));
    expect(child.line.y1).toBe(px(10));
  });
});

describe("resizeChild", () => {
  it("rewrites the transform box", () => {
    const child = { shape: { x: px(0), y: px(0), width: px(10), height: px(10) } };
    resizeChild(child, { x: 5, y: 6, width: 70, height: 80 });
    expect(child.shape.x).toBe(px(5));
    expect(child.shape.width).toBe(px(70));
  });

  it("keeps a line's corner direction", () => {
    // The segment runs bottom-left → top-right; resizing must preserve that.
    const child = { line: { x1: px(0), y1: px(100), x2: px(50), y2: px(0) } };
    resizeChild(child, { x: 10, y: 20, width: 200, height: 60 });
    expect(child.line.x1).toBe(px(10));
    expect(child.line.y1).toBe(px(80)); // 20 + 60 — still the lower corner
    expect(child.line.x2).toBe(px(210));
    expect(child.line.y2).toBe(px(20));
  });
});
