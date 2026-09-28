import { describe, expect, it } from "vitest";

import { slideFillStyle, textBoxFillStyle } from "./text-overlay";

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

describe("slideFillStyle", () => {
  const fade = {
    kind: "gradient" as const,
    stops: [
      { color: "#FFFFFF", position: 0 },
      { color: "#4472C4", position: 100 },
    ],
  };

  it("anchors a slide's linear gradient to the edit box's slice of it", () => {
    expect(slideFillStyle({ ...fade, angle: 90 }, 1280, 720, { x: 0, y: 360, scale: 1 })).toEqual({
      backgroundColor: "#FFFFFF",
      backgroundImage: "linear-gradient(180.0000deg, #FFFFFF -360.00px, #4472C4 360.00px)",
    });
  });

  it("re-centers a slide's radial gradient on the edit box", () => {
    expect(
      slideFillStyle({ ...fade, path: "circle" }, 1280, 720, { x: 0, y: 0, scale: 1 }),
    ).toEqual({
      backgroundColor: "#FFFFFF",
      backgroundImage:
        "radial-gradient(circle 360.00px at 640.00px 360.00px, #FFFFFF 0.00px, #4472C4 360.00px)",
    });
  });

  it("windows a stretched picture background through the edit box", () => {
    expect(
      slideFillStyle({ kind: "image", src: "data:image/png,x" }, 1280, 720, {
        x: 100,
        y: 50,
        scale: 2,
      }),
    ).toEqual({
      backgroundColor: "#FFFFFF",
      backgroundImage: 'url("data:image/png,x")',
      backgroundPosition: "-200px -100px",
      backgroundRepeat: "no-repeat",
      backgroundSize: "2560px 1440px",
    });
  });
});
