import { describe, expect, it } from "vitest";

import { textBoxFillStyle } from "./text-overlay";

describe("textBoxFillStyle", () => {
  it("keeps solid paint alpha out of the text color", () => {
    expect(textBoxFillStyle("FF0000", 0.25)).toEqual({
      backgroundColor: "#FFBFBF",
    });
  });

  it("translates projected gradients to CSS paints", () => {
    expect(
      textBoxFillStyle({
        type: "linear",
        from: { x: 50, y: 0 },
        to: { x: 50, y: 100 },
        stops: [
          { offset: 0, color: "#FF0000" },
          { offset: 1, color: "#FFFFFF" },
        ],
      }),
    ).toEqual({
      backgroundColor: "#FFFFFF",
      backgroundImage: "linear-gradient(180.0000deg, #FF0000 0.00%, #FFFFFF 100.00%)",
    });
  });

  it("translates projected pattern tiles to CSS paints", () => {
    const style = textBoxFillStyle({
      type: "image",
      url: "data:image/svg+xml,tile",
      mode: "repeat",
      repeat: true,
      offset: { x: 2, y: 3 },
      align: "top-left",
    });
    expect(style.backgroundImage).toBe('url("data:image/svg+xml,tile")');
    expect(style.backgroundPosition).toBe("left 2px top 3px");
    expect(style.backgroundRepeat).toBe("repeat");
  });
});
