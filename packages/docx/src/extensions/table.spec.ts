import { describe, it, expect } from "vitest";

import { parseDocx, renderDocx } from "./table";
import { tableFloatToCss } from "./utils";

// tableFloatToCss maps a floating table's w:tblpPr anchor (twips) to CSS. The
// pipeline smoke suites (tests/docx.ts, tests/html.ts) only check that the
// round-trip doesn't throw, so a wrong float side, offset/fromText margin
// collision, or absolute-anchor position would pass — these assertions pin the
// mapping: float side choice, the offset-on-float-side / gap-on-opposite-side
// separation, and the page/margin-anchor absolute-positioning branch.

describe("tableFloatToCss", () => {
  it("float:right — offset on margin-right, leftFromText on margin-left (no clash)", () => {
    const css = tableFloatToCss({
      relativeHorizontalPosition: "right",
      absoluteHorizontalPosition: 720, // → 36pt
      leftFromText: 180, // → 9pt
    });
    expect(css).toEqual(
      expect.arrayContaining(["float:right", "margin-right:36pt", "margin-left:9pt"]),
    );
  });

  it("float:left mirrors — rightFromText → margin-right, no margin-left emitted", () => {
    const css = tableFloatToCss({
      relativeHorizontalPosition: "left",
      rightFromText: 180, // → 9pt
    });
    expect(css).toContain("float:left");
    expect(css).toContain("margin-right:9pt");
    expect(css.some((s) => s.startsWith("margin-left"))).toBe(false);
  });

  it("page anchor + center → absolute, centered horizontally, top at the vpos offset", () => {
    const css = tableFloatToCss({
      horizontalAnchor: "page",
      verticalAnchor: "page",
      relativeHorizontalPosition: "center",
      absoluteVerticalPosition: 1440, // 1" → 72pt
    });
    expect(css).toEqual([
      "position:absolute",
      "top:72pt",
      "left:50%",
      "transform:translateX(-50%)",
    ]);
  });

  it("page/margin anchor without alignment → absolute pinned to left edge", () => {
    // page anchor + relative left (no hpos offset) → best-effort left:0
    expect(
      tableFloatToCss({ horizontalAnchor: "page", relativeHorizontalPosition: "left" }),
    ).toEqual(["position:absolute", "left:0"]);
    // verticalAnchor only (no vpos offset) → just position:absolute
    expect(
      tableFloatToCss({ verticalAnchor: "margin", relativeHorizontalPosition: "left" }),
    ).toEqual(["position:absolute"]);
  });

  it("text-anchored center/inside/outside degrade to [] (no CSS float equivalent)", () => {
    expect(tableFloatToCss({ relativeHorizontalPosition: "center" })).toEqual([]);
    expect(tableFloatToCss({ relativeHorizontalPosition: "inside" })).toEqual([]);
    expect(tableFloatToCss({ relativeHorizontalPosition: "outside" })).toEqual([]);
  });

  it("no relativeHorizontalPosition defaults to float:left", () => {
    expect(tableFloatToCss({})).toContain("float:left");
  });

  it("null/undefined float → []", () => {
    expect(tableFloatToCss(null)).toEqual([]);
    expect(tableFloatToCss(undefined)).toEqual([]);
  });

  it("overlap is ignored (no CSS float equivalent)", () => {
    expect(
      tableFloatToCss({ relativeHorizontalPosition: "left", overlap: "neverOverlap" }),
    ).toEqual(["float:left"]);
  });

  it("top/bottom fromText → margin-top/bottom", () => {
    const css = tableFloatToCss({
      relativeHorizontalPosition: "left",
      topFromText: 200, // → 10pt
      bottomFromText: 200,
    });
    expect(css).toContain("margin-top:10pt");
    expect(css).toContain("margin-bottom:10pt");
  });
});

// float round-trips byte-faithful through the table's renderDocx/parseDocx
// attrs passthrough (SKIP_KEYS drops only rows/columnWidthsRevision) — this is
// what lets a floating table keep its anchor across DOCX→JSON→DOCX even when
// the v1 renderer degrades (absolute anchor, center) and emits no CSS.
describe("table float round-trip", () => {
  it("preserves a full float anchor verbatim", () => {
    const float = {
      horizontalAnchor: "text",
      verticalAnchor: "text",
      absoluteHorizontalPosition: 720,
      absoluteVerticalPosition: 360,
      relativeHorizontalPosition: "right",
      relativeVerticalPosition: "top",
      leftFromText: 180,
      rightFromText: 180,
      topFromText: 200,
      bottomFromText: 200,
      overlap: "neverOverlap",
    };
    const back = parseDocx(renderDocx({ type: "table", attrs: { float } }));
    expect(back.float).toEqual(float);
  });
});
