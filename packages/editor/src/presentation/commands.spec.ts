import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild } from "@docen/pptx";
import type { TextBodyOptions } from "@office-open/core/drawing";
// @vitest-environment node
import { describe, expect, it } from "vitest";

import {
  firstRunSizeOf,
  insertSlideAt,
  makePicture,
  makeTextBox,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  shapeTextOf,
  toggleRunFlag,
  toggleRunStyle,
  writeShapeText,
} from "./commands";

const bodyOf = (paragraphs: unknown[]): TextBodyOptions =>
  ({ paragraphs }) as unknown as TextBodyOptions;

const shapeChild = (body: unknown): SlideChild =>
  ({ shape: { textBody: body } }) as unknown as SlideChild;

type ShapeVariant = Extract<SlideChild, { shape: unknown }>["shape"];
type PictureVariant = Extract<SlideChild, { picture: unknown }>["picture"];

describe("toggleRunFlag", () => {
  it("applies the flag to every run", () => {
    const body = bodyOf([{ children: [{ text: "a" }, { text: "b", bold: true }] }]);
    toggleRunFlag(body, "bold");
    const runs = body.paragraphs as { children: { text: string; bold?: boolean }[] }[];
    expect(runs[0]!.children.map((r) => r.bold)).toEqual([true, true]);
  });

  it("clears the flag when every run carries it", () => {
    const body = bodyOf([
      {
        children: [
          { text: "a", bold: true },
          { text: "b", bold: true },
        ],
      },
    ]);
    toggleRunFlag(body, "bold");
    const runs = body.paragraphs as { children: { bold?: boolean }[] }[];
    expect(runs[0]!.children.map((r) => r.bold)).toEqual([undefined, undefined]);
  });

  it("works through string sugar", () => {
    const body = { text: "hi" } as unknown as TextBodyOptions;
    toggleRunFlag(body, "italic");
    const runs = (body.paragraphs as { children: { text: string; italic?: boolean }[] }[])[0]!
      .children;
    expect(runs).toEqual([{ text: "hi", italic: true }]);
  });

  it("leaves an empty body unchanged", () => {
    const body = bodyOf([{ children: [] }]);
    toggleRunFlag(body, "bold");
    expect((body.paragraphs as { children: unknown[] }[])[0]!.children).toEqual([]);
  });
});

describe("toggleRunStyle", () => {
  it("toggles underline on and back off", () => {
    const body = bodyOf([{ children: [{ text: "a" }] }]);
    toggleRunStyle(body, "underline", "single");
    const run = () =>
      (body.paragraphs as { children: { underline?: string }[] }[])[0]!.children[0]!;
    expect(run().underline).toBe("single");
    toggleRunStyle(body, "underline", "single");
    expect(run().underline).toBeUndefined();
  });

  it("treats the explicit off token as not applied", () => {
    const body = bodyOf([{ children: [{ text: "a", strike: "noStrike" }] }]);
    toggleRunStyle(body, "strike", "singleStrike");
    const run = (body.paragraphs as { children: { strike?: string }[] }[])[0]!.children[0]!;
    expect(run.strike).toBe("singleStrike");
  });
});

describe("run setters", () => {
  it("sets font and size on every run", () => {
    const body = bodyOf([
      { children: [{ text: "a" }] },
      { children: [{ text: "b" }, { text: "c" }] },
    ]);
    setRunFont(body, "Arial");
    setRunSize(body, 18);
    const runs = body.paragraphs as { children: { font?: string; size?: number }[] }[];
    expect(runs.map((p) => p.children.map((r) => [r.font, r.size]))).toEqual([
      [["Arial", 18]],
      [
        ["Arial", 18],
        ["Arial", 18],
      ],
    ]);
  });
});

describe("setParagraphAlignment", () => {
  it("writes the alignment into each paragraph's properties", () => {
    const body = bodyOf([{ children: [{ text: "a" }] }, { properties: { alignment: "left" } }]);
    setParagraphAlignment(body, "center");
    const paragraphs = body.paragraphs as { properties?: { alignment?: string } }[];
    expect(paragraphs.map((p) => p.properties?.alignment)).toEqual(["center", "center"]);
  });

  it("keeps existing paragraph properties", () => {
    const body = bodyOf([{ properties: { level: 2 } }]);
    setParagraphAlignment(body, "right");
    const paragraph = (body.paragraphs as { properties?: Record<string, unknown> }[])[0]!;
    expect(paragraph.properties).toMatchObject({ level: 2, alignment: "right" });
  });
});

describe("insertSlideAt", () => {
  it("splices a blank slide after the index", () => {
    const deck: PresentationOptions = { slides: [{ children: [] }, { children: [] }] };
    insertSlideAt(deck, 0);
    expect(deck.slides).toHaveLength(3);
    expect(deck.slides![1]!.children).toEqual([]);
  });

  it("creates the slides array when missing", () => {
    const deck: PresentationOptions = {};
    insertSlideAt(deck, 0);
    expect(deck.slides).toHaveLength(1);
  });
});

describe("shapeTextOf / writeShapeText", () => {
  it("reads paragraphs as \\n-joined lines", () => {
    const child = shapeChild(
      bodyOf([{ children: [{ text: "a" }] }, { children: [{ text: "b" }] }]),
    );
    expect(shapeTextOf(child)).toBe("a\nb");
  });

  it("reads through the text and string sugar", () => {
    expect(shapeTextOf(shapeChild({ text: "hi" }))).toBe("hi");
    expect(shapeTextOf(shapeChild(bodyOf(["one", "two"])))).toBe("one\ntwo");
  });

  it("returns null for children without a text body", () => {
    const { picture } = makePicture(100, 100, 10, 10, "data:", "png") as {
      picture: PictureVariant;
    };
    expect(shapeTextOf({ picture } as unknown as SlideChild)).toBeNull();
  });

  it("writes one paragraph per line, keeping alignment and first-run style", () => {
    const child = shapeChild(
      bodyOf([
        { properties: { alignment: "center" }, children: [{ text: "old", bold: true, size: 24 }] },
      ]),
    );
    expect(writeShapeText(child, "one\ntwo")).toBe(true);
    const body = (child as { shape: { textBody: TextBodyOptions } }).shape.textBody;
    const paragraphs = body.paragraphs as {
      properties?: { alignment?: string };
      children: { text: string; bold?: boolean; size?: number }[];
    }[];
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]!.properties?.alignment).toBe("center");
    expect(paragraphs[0]!.children[0]).toMatchObject({ text: "one", bold: true, size: 24 });
    // A paragraph beyond the originals carries no inherited style.
    expect(paragraphs[1]!.children[0]).toEqual({ text: "two" });
  });

  it("clears the text sugar after writing", () => {
    const child = shapeChild({ text: "before" });
    writeShapeText(child, "after");
    const body = (child as { shape: { textBody: TextBodyOptions } }).shape.textBody;
    expect(body.text).toBeUndefined();
    expect(shapeTextOf(child)).toBe("after");
  });
});

describe("firstRunSizeOf", () => {
  it("returns the first run's size", () => {
    const child = shapeChild(bodyOf([{ children: [{ text: "a", size: 32 }] }]));
    expect(firstRunSizeOf(child)).toBe(32);
  });

  it("falls back to 18pt when unset or absent", () => {
    expect(firstRunSizeOf(shapeChild(bodyOf([{ children: [{ text: "a" }] }])))).toBe(18);
    expect(firstRunSizeOf(makeTextBox(100, 100))).toBe(18);
  });
});

describe("makeTextBox", () => {
  it("centers a 40%×15% box with a visible outline", () => {
    const { shape } = makeTextBox(1280, 720) as { shape: ShapeVariant };
    expect(shape.x).toBe(Math.round(1280 * 0.3 * EMU_PER_PX));
    expect(shape.y).toBe(Math.round(720 * 0.425 * EMU_PER_PX));
    expect(shape.width).toBe(Math.round(1280 * 0.4 * EMU_PER_PX));
    expect(shape.height).toBe(Math.round(720 * 0.15 * EMU_PER_PX));
    expect(shape.properties).toMatchObject({
      geometry: "rect",
      fill: "FFFFFF",
      outline: { width: "1pt" },
    });
  });
});

describe("makePicture", () => {
  it("scales a large image into half the slide, centered", () => {
    const { picture } = makePicture(1280, 720, 2560, 1440, "data:", "png") as {
      picture: PictureVariant;
    };
    expect(picture.width).toBe(Math.round(640 * EMU_PER_PX));
    expect(picture.height).toBe(Math.round(360 * EMU_PER_PX));
    expect(picture.x).toBe(Math.round(320 * EMU_PER_PX));
    expect(picture.y).toBe(Math.round(180 * EMU_PER_PX));
  });

  it("never upscales a small image", () => {
    const { picture } = makePicture(1280, 720, 100, 80, "data:", "png") as {
      picture: PictureVariant;
    };
    expect(picture.width).toBe(100 * EMU_PER_PX);
    expect(picture.height).toBe(80 * EMU_PER_PX);
  });
});
