import { describe, expect, it } from "vitest";

import { outlineLineOf } from "./shape-fill";

describe("PPTX outline fills", () => {
  it("maps gradient strokes to renderer-native linear paint", () => {
    const line = outlineLineOf(
      {
        width: 12700,
        type: "gradFill",
        gradientFill: {
          stops: [
            { position: 0, color: "FF0000" },
            { position: 100, color: "0000FF" },
          ],
          shade: { angle: 0 },
        },
      },
      100,
      50,
    );
    expect(line?.stroke).toEqual({
      type: "linear",
      from: { x: 0, y: 25 },
      to: { x: 100, y: 25 },
      stops: [
        { offset: 0, color: "#FF0000" },
        { offset: 1, color: "#0000FF" },
      ],
    });
  });

  it("maps pattern strokes to renderer-native repeated image paint", () => {
    const line = outlineLineOf(
      {
        width: 25400,
        type: "pattFill",
        patternFill: {
          pattern: "cross",
          foregroundColor: { value: "FF0000" },
          backgroundColor: { value: "FFFFFF" },
        },
      },
      80,
      60,
    );
    expect(line?.stroke).toMatchObject({ type: "image", mode: "repeat", repeat: true });
  });

  it("keeps solid stroke alpha in the paint", () => {
    const line = outlineLineOf(
      {
        width: 12700,
        type: "solidFill",
        color: { value: "FF0000", transforms: { alpha: 50 } },
      },
      40,
      30,
    );
    expect(line?.stroke).toBe("rgba(255, 0, 0, 0.5)");
  });
});
