import { describe, expect, it } from "vitest";

import { browserFontMetrics, clearFontMetricCache } from "./font";
import { WORD_FONT_METRICS, wordLineRatio } from "./font-metrics-data";

describe("word font metrics table", () => {
  it("computes Word's formula ratio from the OS/2 tables", () => {
    // SimSun's published Word ratio: (220 + 36 + 2×38) / 256 = 1.296875.
    expect(wordLineRatio({ upem: 256, winAscent: 220, winDescent: 36 })).toBe(1.296875);
    expect(wordLineRatio({ upem: 2048, winAscent: 1854, winDescent: 434 })).toBeCloseTo(
      1.452148,
      6,
    );
  });

  it("resolves localized family aliases to the same face", () => {
    expect(WORD_FONT_METRICS["宋体"]).toBe(WORD_FONT_METRICS.simsun);
    expect(WORD_FONT_METRICS["微软雅黑"]).toEqual(WORD_FONT_METRICS["microsoft yahei"]);
  });

  it("serves tabulated faces Word's ratio without probing", () => {
    clearFontMetricCache();
    // Node has no DOM: the probe path returns the 1.2 fallback, so a
    // non-1.2 answer proves the table was consulted.
    expect(browserFontMetrics.normalRatio({ family: "SimSun" })).toBe(1.296875);
    expect(browserFontMetrics.normalRatio({ family: "宋体", bold: true })).toBe(1.296875);
    expect(browserFontMetrics.normalRatio({ family: " Arial " })).toBeCloseTo(1.452148, 6);
  });

  it("falls back to the probe ratio for untabulated families", () => {
    clearFontMetricCache();
    expect(browserFontMetrics.normalRatio({ family: "NotARealFont" })).toBe(1.2);
  });
});
