import { describe, expect, it } from "vitest";

import { resolveChartColors } from "./chart-colors";

describe("resolveChartColors", () => {
  it("maps DrawingML aliases through the Office color map", () => {
    const resolved = resolveChartColors(
      {
        background: { value: "bg1" },
        foreground: { value: "tx1" },
        hyperlink: { value: "hlink" },
      },
      {
        themeColors: {
          light1: "FFFFFF",
          dark1: "111111",
          hyperlink: "222222",
        },
      },
    );

    expect(resolved).toEqual({
      background: "FFFFFF",
      foreground: "111111",
      hyperlink: "222222",
    });
  });

  it("honors an override before the default slot mapping", () => {
    const resolved = resolveChartColors(
      { fill: { value: "bg1" } },
      {
        themeColors: { dark1: "111111", light1: "FFFFFF" },
        colorMapping: { bg1: "tx1" },
      },
    );

    expect(resolved).toEqual({ fill: "111111" });
  });
});
