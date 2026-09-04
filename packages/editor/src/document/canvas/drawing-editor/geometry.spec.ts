import { describe, expect, it } from "vitest";

import { cropFullBox, handleAt, resizeBox, resizeCrop, rotateDelta, type Box } from "./geometry";

const box: Box = { x: 100, y: 80, width: 200, height: 100 };

describe("handleAt", () => {
  it("finds corner and edge handles on the frame", () => {
    expect(handleAt(box, 100, 80)).toBe("nw");
    expect(handleAt(box, 300, 180)).toBe("se");
    expect(handleAt(box, 200, 80)).toBe("n");
    expect(handleAt(box, 300, 130)).toBe("e");
  });

  it("returns null inside the body and outside the frame", () => {
    expect(handleAt(box, 200, 130)).toBeNull();
    expect(handleAt(box, 50, 50)).toBeNull();
  });

  it("prefers a corner when the pointer sits in both grab zones", () => {
    // 3px off the NW corner is within the n, w, and nw grab zones.
    expect(handleAt(box, 97, 77)).toBe("nw");
  });
});

describe("resizeBox", () => {
  it("anchors the opposite corner on a corner drag", () => {
    const next = resizeBox(box, "se", 50, 20); // width drives, height follows
    expect(next).toEqual({ x: 100, y: 80, width: 250, height: 125 });
    const grown = resizeBox(box, "nw", -30, -10); // x/y move, anchor (se) fixed
    expect(grown).toEqual({ x: 70, y: 65, width: 230, height: 115 });
  });

  it("keeps the aspect ratio on corner drags (Word's picture default)", () => {
    const next = resizeBox(box, "se", 100, 0); // width drives
    expect(next.width / next.height).toBeCloseTo(2);
    expect(next.height).toBeCloseTo(150);
  });

  it("resizes one axis freely from an edge handle", () => {
    const next = resizeBox(box, "e", 50, 999);
    expect(next).toEqual({ x: 100, y: 80, width: 250, height: 100 });
  });

  it("clamps to the minimum instead of inverting past the anchor", () => {
    const next = resizeBox(box, "se", -500, -500);
    expect(next.width).toBe(24);
    expect(next.height).toBe(24);
    expect(next.x).toBe(100);
    expect(next.y).toBe(80);
  });

  it("moves the anchored edge when dragging the west/north below minimum", () => {
    const next = resizeBox(box, "nw", 500, 500);
    expect(next.x).toBe(300 - 24);
    expect(next.y).toBe(180 - 24);
    expect(next.width).toBe(24);
    expect(next.height).toBe(24);
  });
});

describe("rotateDelta", () => {
  const c = 0; // center at the origin

  it("measures the clockwise sweep between two pointer positions", () => {
    // Screen px: +y is DOWN, so right→below sweeps +90° (clockwise).
    expect(rotateDelta(c, c, 50, 0, 0, 50)).toBe(90);
    expect(rotateDelta(c, c, 50, 0, 0, -50)).toBe(-90);
    expect(rotateDelta(c, c, 50, 0, 50, 0)).toBe(0);
  });

  it("keeps the spin continuous across the ±180° wrap", () => {
    // Left of the center is atan2's 180° boundary; stepping to the top is a
    // +90° clockwise sweep, not a -270° rewind.
    expect(rotateDelta(c, c, -50, 0, 0, -50)).toBe(90);
    // And counter-clockwise over the same boundary stays negative.
    expect(rotateDelta(c, c, 0, -50, -50, 0)).toBe(-90);
  });
});

describe("cropFullBox", () => {
  it("expands the visible box to the full source", () => {
    // Visible 100×50 covers the middle half of each axis (0.25 in from each
    // side) — the source is double the size, shifted by the crop's share.
    const full = cropFullBox(
      { x: 100, y: 100, width: 100, height: 50 },
      {
        left: 0.25,
        top: 0.25,
        right: 0.25,
        bottom: 0.25,
      },
    );
    expect(full).toEqual({ x: 50, y: 75, width: 200, height: 100 });
  });

  it("is the visible box itself on an empty crop", () => {
    expect(cropFullBox(box, { left: 0, top: 0, right: 0, bottom: 0 })).toEqual(box);
  });
});

describe("resizeCrop", () => {
  const crop = { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 };

  it("grows the inset whose edge drags inward", () => {
    // West handle rightward grows `left`; east handle rightward shrinks
    // `right` (the right edge follows the pointer, leaving more source).
    expect(resizeCrop(crop, "w", 0.1, 0, 0, 0).left).toBe(0.2);
    expect(resizeCrop(crop, "e", 0.05, 0, 0.05, 0).right).toBe(0.05);
    expect(resizeCrop(crop, "n", 0, 0.15, 0, 0).top).toBe(0.25);
    expect(resizeCrop(crop, "s", 0, 0, 0, -0.05).bottom).toBeCloseTo(0.15);
  });

  it("clamps so an axis never crops past a visible remainder", () => {
    const next = resizeCrop(crop, "w", 5, 0, 0, 0);
    expect(next.left + next.right).toBeLessThanOrEqual(0.9);
    expect(next.left).toBeGreaterThanOrEqual(0);
  });
});
