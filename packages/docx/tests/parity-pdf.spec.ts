import { PNG } from "pngjs";
import { describe, expect, it } from "vitest";

import { comparePng, DEFAULT_CHANNEL_TOLERANCE } from "./parity-pdf";

function png(colors: readonly [number, number, number][]): Uint8Array {
  const image = new PNG({ width: colors.length, height: 1 });
  colors.forEach(([red, green, blue], index) => {
    const offset = index * 4;
    image.data[offset] = red;
    image.data[offset + 1] = green;
    image.data[offset + 2] = blue;
    image.data[offset + 3] = 255;
  });
  return new Uint8Array(PNG.sync.write(image));
}

describe("PDF parity pixel comparator", () => {
  it("counts only pixels beyond the antialiasing tolerance", () => {
    const expected = png([
      [0, 0, 0],
      [255, 255, 255],
    ]);
    const actual = png([
      [DEFAULT_CHANNEL_TOLERANCE, DEFAULT_CHANNEL_TOLERANCE, DEFAULT_CHANNEL_TOLERANCE],
      [255 - DEFAULT_CHANNEL_TOLERANCE - 1, 255, 255],
    ]);

    expect(comparePng(expected, actual)).toMatchObject({
      page: 1,
      width: 2,
      height: 1,
      changedPixels: 1,
      changedRatio: 0.5,
    });
  });

  it("fails dimension drift before comparing pixel data", () => {
    expect(
      comparePng(
        png([[0, 0, 0]]),
        png([
          [0, 0, 0],
          [255, 255, 255],
        ]),
      ),
    ).toMatchObject({
      width: 2,
      changedRatio: 1,
      reason: "dimension-mismatch",
    });
  });
});
