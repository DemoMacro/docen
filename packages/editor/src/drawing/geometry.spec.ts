import { describe, expect, it } from "vitest";

import {
  cropFullBox,
  freshChildEmu,
  handleAt,
  memberEmuOf,
  resizeBox,
  resizeCrop,
  rotateDelta,
  unionBox,
  type Box,
} from "./geometry";

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

  it("allows coordinates to reach zero or negative values without clamping to 1", () => {
    const next = resizeBox({ x: 10, y: 10, width: 50, height: 50 }, "w", -20, 0);
    expect(next.x).toBe(-10);
    expect(next.width).toBe(70);
  });

  it("keeps the ratio from the dragged edge's midpoint when locked (Shift)", () => {
    // East drag: width drives, height follows, centered on the west edge.
    const east = resizeBox(box, "e", 50, 0, 24, true);
    expect(east).toEqual({ x: 75, y: 68, width: 250, height: 125 });
    // North drag: height drives, width follows, centered on the south edge.
    const north = resizeBox(box, "n", 0, -20, 24, true);
    expect(north).toEqual({ x: 80, y: 70, width: 240, height: 120 });
  });

  it("clamps a locked edge drag to the minimum without breaking the ratio", () => {
    const next = resizeBox(box, "s", 0, -500, 24, true);
    expect(next.width / next.height).toBeCloseTo(2);
    expect(next.height).toBe(24);
    expect(next.x).toBe(100 + (200 - 48) / 2);
    expect(next.y).toBe(80 + (100 - 24) / 2);
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

describe("unionBox", () => {
  it("bounds every member", () => {
    expect(
      unionBox([
        { x: 100, y: 80, width: 200, height: 100 },
        { x: 50, y: 150, width: 120, height: 60 },
      ]),
    ).toEqual({ x: 50, y: 80, width: 250, height: 130 });
  });
});

describe("freshChildEmu / memberEmuOf round-trip", () => {
  const members: Box[] = [
    { x: 400, y: 300, width: 200, height: 100 },
    { x: 350, y: 350, width: 120, height: 60 },
  ];
  const union = unionBox(members);
  const EMU = 9525;

  it("maps each member to its EMU offset from the union's top-left", () => {
    expect(freshChildEmu(members[0]!, union)).toEqual({
      x: 50 * EMU,
      y: 0,
      cx: 200 * EMU,
      cy: 100 * EMU,
    });
    expect(freshChildEmu(members[1]!, union)).toEqual({
      x: 0,
      y: 50 * EMU,
      cx: 120 * EMU,
      cy: 60 * EMU,
    });
  });

  it("reverses exactly on a fresh 1:1 group (chOff 0, chExt = ext)", () => {
    // Group the members, then ungroup: every child comes back to its page
    // position within 1 EMU of rounding.
    const ext = { x: union.width * EMU, y: union.height * EMU };
    for (const m of members) {
      const child = freshChildEmu(m, union);
      const back = memberEmuOf(child, { x: 0, y: 0 }, ext, ext);
      expect(Math.abs(back.dx - (m.x - union.x) * EMU)).toBeLessThanOrEqual(1);
      expect(Math.abs(back.dy - (m.y - union.y) * EMU)).toBeLessThanOrEqual(1);
      expect(Math.abs(back.cx - m.width * EMU)).toBeLessThanOrEqual(1);
      expect(Math.abs(back.cy - m.height * EMU)).toBeLessThanOrEqual(1);
    }
  });

  it("scales through a non-1:1 child space (ext ≠ chExt)", () => {
    // A Word-authored group with chExt half its display extent: child EMU
    // covers the same area at half the page scale.
    const back = memberEmuOf(
      { x: 100000, y: 0, cx: 95250, cy: 95250 },
      { x: 50000, y: 0 },
      { x: 190500, y: 190500 },
      { x: 95250, y: 95250 },
    );
    expect(back).toEqual({ dx: 100000, dy: 0, cx: 190500, cy: 190500 });
  });

  it("treats a missing chExt as 1:1 (the projection's childScale convention)", () => {
    const back = memberEmuOf(
      { x: 100000, y: 50000, cx: 9525, cy: 9525 },
      { x: 100000, y: 50000 },
      { x: 190500, y: 190500 },
      undefined,
    );
    expect(back).toEqual({ dx: 0, dy: 0, cx: 9525, cy: 9525 });
  });
});
