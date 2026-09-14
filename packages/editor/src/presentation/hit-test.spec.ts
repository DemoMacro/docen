import { EMU_PER_PX } from "@docen/layout";
import type { SlideOptions } from "@docen/pptx";
// @vitest-environment node
import { describe, expect, it } from "vitest";

import {
  captureGeometry,
  hitSlide,
  offsetChild,
  resizeChild,
  restoreGeometry,
  rotateChild,
  slideHits,
} from "./hit-test";

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
        { table: { x: px(1), y: px(2), width: px(60), height: px(30), rows: [] } },
        // Unpainted variants contribute no hit box.
        { smartart: {} },
      ],
    } as unknown as SlideOptions;
    const hits = slideHits(slide);
    expect(hits.map((h) => h.child)).toEqual([0, 1, 2, 3]);
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

describe("rotateChild", () => {
  it("accumulates a transform child's spin", () => {
    const child = { shape: { x: px(0), y: px(0), width: px(10), height: px(10), rotation: 30 } };
    rotateChild(child, 15);
    expect(child.shape.rotation).toBe(45);
    rotateChild(child, -100);
    expect(child.shape.rotation).toBe(-55);
  });

  it("spins a line's endpoints around the box center", () => {
    // A horizontal segment through its bounding box center: a 180° sweep
    // swaps the endpoints.
    const child = { line: { x1: px(0), y1: px(10), x2: px(100), y2: px(10) } };
    rotateChild(child, 180);
    expect(child.line.x1).toBe(px(100));
    expect(child.line.y1).toBe(px(10));
    expect(child.line.x2).toBe(px(0));
    expect(child.line.y2).toBe(px(10));
  });

  it("never spins the spin-less frames", () => {
    // Table/chart frames carry no a:xfrm @rot — the gesture must not invent
    // the field.
    const child = { table: { x: px(0), y: px(0), width: px(10), height: px(10), rows: [] } };
    rotateChild(child, 45);
    expect("rotation" in child.table).toBe(false);
  });

  it("captures and restores the spin", () => {
    const child = { shape: { x: px(0), y: px(0), width: px(10), height: px(10), rotation: 45 } };
    const snap = captureGeometry(child);
    rotateChild(child, 15);
    restoreGeometry(child, snap);
    expect(child.shape.rotation).toBe(45);
  });
});
