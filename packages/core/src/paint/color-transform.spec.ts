import { describe, expect, it } from "vitest";

import { transformColor } from "./color-transform";

describe("DrawingML color transforms", () => {
  it("applies tint, shade, and source-order composition", () => {
    expect(transformColor("FF0000", { tint: 40 })).toEqual({ color: "FFCBCB", alpha: 1 });
    expect(transformColor("FF0000", { shade: 60 })).toEqual({ color: "CB0000", alpha: 1 });
    // 40% tint followed by 50% shade differs from the reverse composition.
    const tintShade = transformColor("FF0000", { tint: 40, shade: 50 }).color;
    const shadeTint = transformColor("FF0000", { shade: 50, tint: 40 }).color;
    expect(tintShade).toBe("BC9595");
    expect(shadeTint).toBe("E7CBCB");
  });

  it("evaluates hue, saturation, luminance, channel, and switch transforms", () => {
    expect(transformColor("FF0000", { hueOff: 120 }).color).toBe("00FF00");
    expect(transformColor("FF0000", { satMod: 50 }).color).toBe("BF4040");
    expect(transformColor("FF0000", { lumMod: 20, lumOff: 80 }).color).toBe("FFCCCC");
    expect(transformColor("336699", { redOff: 20 }).color).toBe("666699");
    expect(transformColor("FF0000", { inv: true }).color).toBe("00FFFF");
    expect(transformColor("FF0000", { comp: true }).color).toBe("00FFFF");
    expect(transformColor("FF0000", { gray: true }).color).toBe("4C4C4C");
  });

  it("keeps the alpha channel separate from RGB", () => {
    expect(transformColor("FF0000", { alpha: 50, alphaMod: 80 })).toEqual({
      color: "FF0000",
      alpha: 0.4,
    });
  });
});
