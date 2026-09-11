import { beforeEach, describe, expect, it } from "vitest";

import { installFakeCanvas } from "../test/fake-canvas";
import {
  clearMeasurementCaches,
  getCorrectedSegmentWidth,
  getFontMeasurementState,
  getSegmentBreakableFitAdvances,
  getSegmentMetricCache,
  getSegmentMetrics,
} from "./measurement";
import {
  layoutNextRichInlineLineRange,
  materializeRichInlineLineRange,
  prepareRichInline,
  walkRichInlineLineRanges,
} from "./rich-inline";

installFakeCanvas();

const FONT = "16px serif";

beforeEach(() => {
  // The measurement module keeps ONE canvas context whose `.font` persists
  // across specs; re-pin it to the suite font and drop cached segment metrics.
  clearMeasurementCaches();
  getFontMeasurementState(FONT, false, false);
});

describe.sequential("vendored CJK correction", () => {
  it("gates the DOM probe off without a document and reports zero correction", () => {
    // The fake's CJK advance rounds 14.67px up to 15 (> fontSize + 0.01), so
    // the gate opens; with no usable `document` the probe is skipped and the
    // correction stays 0 instead of throwing.
    const state = getFontMeasurementState("14.67px serif", false, true);
    expect(state.cjkCorrection).toBe(0);
  });

  it("subtracts per-glyph correction in getCorrectedSegmentWidth", () => {
    const cache = getSegmentMetricCache(FONT);
    const metrics = getSegmentMetrics("甲乙", cache);
    expect(metrics.width).toBe(32);
    expect(getCorrectedSegmentWidth("甲乙", metrics, 0, 0.33)).toBeCloseTo(32 - 2 * 0.33, 10);
  });

  it("passes the correction through breakable fit advances", () => {
    const cache = getSegmentMetricCache(FONT);
    const metrics = getSegmentMetrics("甲乙", cache);
    const advances = getSegmentBreakableFitAdvances(
      "甲乙",
      metrics,
      cache,
      0,
      0.33,
      "sum-graphemes",
    );
    expect(advances).not.toBeNull();
    expect(advances![0]).toBeCloseTo(16 - 0.33, 10);
  });
});

describe.sequential("vendored empty-text atom retention", () => {
  it("keeps a zero-text extraWidth atom as an unbreakable fragment", () => {
    // An inline image rides as { text: "", extraWidth }: upstream collapsed it
    // away entirely, the vendored fix keeps it as a break:'never' item.
    const prepared = prepareRichInline([
      { text: "甲乙", font: FONT },
      { text: "", font: FONT, extraWidth: 30 },
    ]);
    const range = layoutNextRichInlineLineRange(prepared, 200);
    expect(range).not.toBeNull();
    const line = materializeRichInlineLineRange(prepared, range!);
    expect(line.fragments).toHaveLength(2);
    expect(line.fragments[1]!.text).toBe("");
    expect(line.width).toBe(62); // 32 text + 30 extra
  });

  it("still collapses whitespace-only items without extraWidth", () => {
    const prepared = prepareRichInline([{ text: "  ", font: FONT }]);
    expect(layoutNextRichInlineLineRange(prepared, 200)).toBeNull();
  });

  it("credits a collapsed space ahead of the atom to its gapBefore", () => {
    const prepared = prepareRichInline([
      { text: "甲", font: FONT },
      { text: " ", font: FONT },
      { text: "", font: FONT, extraWidth: 30 },
    ]);
    const range = layoutNextRichInlineLineRange(prepared, 200);
    const line = materializeRichInlineLineRange(prepared, range!);
    expect(line.fragments).toHaveLength(2);
    expect(line.fragments[1]!.gapBefore).toBe(4); // 16px em / 4
    expect(line.width).toBe(50); // 16 text + 4 gap + 30 extra
  });
});

describe.sequential("rich inline pre-wrap (preserved spaces)", () => {
  // The fake's advances at 16px: CJK 16, latin 8, space 4.
  it("keeps every space of a run as its own advance", () => {
    const prepared = prepareRichInline([{ text: "a  b", font: FONT }], {
      whiteSpace: "pre-wrap",
    });
    const range = layoutNextRichInlineLineRange(prepared, 200);
    const line = materializeRichInlineLineRange(prepared, range!);
    // "a  b" = 8 + 4 + 4 + 8 — both spaces paid, none collapsed into a gap.
    expect(line.fragments).toHaveLength(1);
    expect(line.fragments[0]!.gapBefore).toBe(0);
    expect(line.fragments[0]!.text).toBe("a  b");
    expect(line.width).toBe(24);
  });

  it("lays a whitespace-only item as a real text fragment", () => {
    const prepared = prepareRichInline([{ text: "  ", font: FONT }], {
      whiteSpace: "pre-wrap",
    });
    const range = layoutNextRichInlineLineRange(prepared, 200);
    const line = materializeRichInlineLineRange(prepared, range!);
    expect(line.fragments).toHaveLength(1);
    expect(line.fragments[0]!.text).toBe("  ");
    expect(line.width).toBe(8);
  });

  it("keeps the inter-item space inside the item's text, not a gap", () => {
    const prepared = prepareRichInline(
      [
        { text: "甲", font: FONT },
        { text: " 乙", font: FONT },
      ],
      { whiteSpace: "pre-wrap" },
    );
    const range = layoutNextRichInlineLineRange(prepared, 200);
    const line = materializeRichInlineLineRange(prepared, range!);
    expect(line.fragments).toHaveLength(2);
    expect(line.fragments[1]!.text).toBe(" 乙");
    expect(line.fragments[1]!.gapBefore).toBe(0);
    expect(line.width).toBe(36); // 16 + 4 + 16 (CJK)
  });

  it("hangs line-trailing preserved spaces past the wrap width", () => {
    // maxWidth 18 fits "ab" (16); the trailing spaces exceed it but hang —
    // Word/CSS pre-wrap keeps them on the line instead of pushing them to
    // the next line's start (where they would be consumed in normal mode).
    const prepared = prepareRichInline([{ text: "ab   cd", font: FONT }], {
      whiteSpace: "pre-wrap",
    });
    const line1 = layoutNextRichInlineLineRange(prepared, 18);
    expect(line1).not.toBeNull();
    const l1 = materializeRichInlineLineRange(prepared, line1!);
    expect(l1.fragments[0]!.text).toBe("ab   ");
    const line2 = layoutNextRichInlineLineRange(prepared, 18, line1!.end);
    expect(line2).not.toBeNull();
    const l2 = materializeRichInlineLineRange(prepared, line2!);
    expect(l2.fragments[0]!.text).toBe("cd");
  });

  it("terminates on a longer-than-line space run without looping", () => {
    // A space run is ONE segment and preserved spaces hang past the fit
    // limit, so an over-long run rides one line (known limit — inner breaks
    // are a text-segment feature). The walk must still terminate boundedly.
    const prepared = prepareRichInline([{ text: `a${" ".repeat(200)}`, font: FONT }], {
      whiteSpace: "pre-wrap",
    });
    let lineCount = 0;
    walkRichInlineLineRanges(prepared, 100, () => lineCount++);
    expect(lineCount).toBeGreaterThanOrEqual(1);
    expect(lineCount).toBeLessThan(210);
  });
});
