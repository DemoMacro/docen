import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild, TableCellOptions } from "@docen/pptx";
import type { ParagraphDescriptorOptions, TextBodyOptions } from "@office-open/core/drawing";
// @vitest-environment node
import { describe, expect, it } from "vitest";

import {
  cellTextOf,
  deleteSlideAt,
  duplicateSlideAt,
  firstRunSizeOf,
  insertSlideAt,
  makePicture,
  makeTextBox,
  reorderChild,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  shapeTextOf,
  toggleRunFlag,
  toggleRunStyle,
  writeCellText,
  writeShapeText,
} from "./commands";

type Run = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: string;
  font?: string;
  size?: number;
};
type Para = { properties?: { alignment?: string; level?: number }; children?: (Run | string)[] };
const paras = (paragraphs: Para[]): ParagraphDescriptorOptions[] =>
  paragraphs as unknown as ParagraphDescriptorOptions[];

const bodyOf = (paragraphs: unknown[]): TextBodyOptions =>
  ({ paragraphs }) as unknown as TextBodyOptions;

const shapeChild = (body: unknown): SlideChild =>
  ({ shape: { textBody: body } }) as unknown as SlideChild;

const cellOf = (children: unknown[]): TableCellOptions =>
  ({ children }) as unknown as TableCellOptions;

type ShapeVariant = Extract<SlideChild, { shape: unknown }>["shape"];
type PictureVariant = Extract<SlideChild, { picture: unknown }>["picture"];

describe("toggleRunFlag", () => {
  it("applies the flag to every run", () => {
    const paragraphs = paras([{ children: [{ text: "a" }, { text: "b", bold: true }] }]);
    toggleRunFlag(paragraphs, "bold");
    expect(paragraphs[0]!.children!.map((r) => (r as Run).bold)).toEqual([true, true]);
  });

  it("clears the flag when every run carries it", () => {
    const paragraphs = paras([
      {
        children: [
          { text: "a", bold: true },
          { text: "b", bold: true },
        ],
      },
    ]);
    toggleRunFlag(paragraphs, "bold");
    expect(paragraphs[0]!.children!.map((r) => (r as Run).bold)).toEqual([undefined, undefined]);
  });

  it("works through paragraph text sugar", () => {
    const paragraphs = paras([{ text: "hi" } as unknown as Para]);
    toggleRunFlag(paragraphs, "italic");
    expect(paragraphs[0]!.children).toEqual([{ text: "hi", italic: true }]);
  });

  it("leaves paragraphs without runs unchanged", () => {
    const paragraphs = paras([{ children: [] }]);
    toggleRunFlag(paragraphs, "bold");
    expect(paragraphs[0]!.children).toEqual([]);
  });
});

describe("toggleRunStyle", () => {
  it("toggles underline on and back off", () => {
    const paragraphs = paras([{ children: [{ text: "a" }] }]);
    toggleRunStyle(paragraphs, "underline", "single");
    const run = () => paragraphs[0]!.children![0] as Run;
    expect(run().underline).toBe("single");
    toggleRunStyle(paragraphs, "underline", "single");
    expect(run().underline).toBeUndefined();
  });

  it("treats the explicit off token as not applied", () => {
    const paragraphs = paras([{ children: [{ text: "a", strike: "noStrike" }] }]);
    toggleRunStyle(paragraphs, "strike", "singleStrike");
    expect((paragraphs[0]!.children![0] as Run).strike).toBe("singleStrike");
  });
});

describe("run setters", () => {
  it("sets font and size on every run", () => {
    const paragraphs = paras([
      { children: [{ text: "a" }] },
      { children: [{ text: "b" }, { text: "c" }] },
    ]);
    setRunFont(paragraphs, "Arial");
    setRunSize(paragraphs, 18);
    expect(paragraphs.map((p) => (p.children as Run[]).map((r) => [r.font, r.size]))).toEqual([
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
    const paragraphs = paras([
      { children: [{ text: "a" }] },
      { properties: { alignment: "left" } },
    ] as Para[]);
    setParagraphAlignment(paragraphs, "center");
    expect(paragraphs.map((p) => p.properties?.alignment)).toEqual(["center", "center"]);
  });

  it("keeps existing paragraph properties", () => {
    const paragraphs = paras([{ properties: { level: 2 } }] as Para[]);
    setParagraphAlignment(paragraphs, "right");
    expect(paragraphs[0]!.properties).toMatchObject({ level: 2, alignment: "right" });
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

describe("deleteSlideAt / duplicateSlideAt", () => {
  it("removes the slide and returns it", () => {
    const gone = { children: [] };
    const deck: PresentationOptions = { slides: [{ children: [] }, gone] };
    expect(deleteSlideAt(deck, 1)).toBe(gone);
    expect(deck.slides).toHaveLength(1);
  });

  it("returns null outside the deck", () => {
    const deck: PresentationOptions = { slides: [{ children: [] }] };
    expect(deleteSlideAt(deck, 5)).toBeNull();
  });

  it("duplicates deep — the copy shares no child objects", () => {
    const deck: PresentationOptions = {
      slides: [{ children: [{ shape: { textBody: { text: "hi" } } }] }],
    };
    const copyAt = duplicateSlideAt(deck, 0);
    expect(copyAt).toBe(1);
    expect(deck.slides).toHaveLength(2);
    expect(deck.slides![1]).not.toBe(deck.slides![0]);
    expect(deck.slides![1]!.children![0]).not.toBe(deck.slides![0]!.children![0]);
  });

  it("returns -1 outside the deck", () => {
    expect(duplicateSlideAt({}, 0)).toBe(-1);
  });
});

describe("reorderChild", () => {
  const items = (): SlideChild[] => ["a", "b", "c"] as unknown as SlideChild[];

  it("moves a child to the top (last paints on top)", () => {
    const children = items();
    reorderChild(children, 0, 2);
    expect(children).toEqual(["b", "c", "a"]);
  });

  it("moves a child to the bottom", () => {
    const children = items();
    reorderChild(children, 2, 0);
    expect(children).toEqual(["c", "a", "b"]);
  });

  it("clamps the target index", () => {
    const children = items();
    reorderChild(children, 0, 99);
    expect(children).toEqual(["b", "c", "a"]);
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

  it("reads a soft break as a line separator", () => {
    const child = shapeChild(
      bodyOf([{ children: [{ text: "a" }, { break: true }, { text: "b" }] }]),
    );
    expect(shapeTextOf(child)).toBe("a\nb");
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

describe("cellTextOf / writeCellText", () => {
  it("reads the cell's paragraphs as \\n-joined lines", () => {
    const cell = cellOf([{ children: [{ text: "a" }] }, { children: [{ text: "b" }] }]);
    expect(cellTextOf(cell)).toBe("a\nb");
  });

  it("reads through the text and string sugar", () => {
    expect(cellTextOf({ text: "hi" } as unknown as TableCellOptions)).toBe("hi");
    expect(cellTextOf(cellOf(["one", "two"]))).toBe("one\ntwo");
  });

  it("writes one paragraph per line, keeping alignment and first-run style", () => {
    const cell = cellOf([
      { properties: { alignment: "center" }, children: [{ text: "old", bold: true, size: 24 }] },
    ]);
    writeCellText(cell, "one\ntwo");
    const paragraphs = cell.children as {
      properties?: { alignment?: string };
      children: { text: string; bold?: boolean; size?: number }[];
    }[];
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]!.properties?.alignment).toBe("center");
    expect(paragraphs[0]!.children[0]).toMatchObject({ text: "one", bold: true, size: 24 });
    // A paragraph beyond the originals carries no inherited style.
    expect(paragraphs[1]!.children[0]).toEqual({ text: "two" });
    expect(cell.text).toBeUndefined();
  });

  it("round-trips empty text", () => {
    const cell = cellOf([]);
    writeCellText(cell, "");
    expect(cellTextOf(cell)).toBe("");
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
