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
