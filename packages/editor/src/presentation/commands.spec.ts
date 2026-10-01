import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild, SlideOptions, TableCellOptions } from "@docen/pptx";
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
  makeShape,
  makeTable,
  makeTextBox,
  makeFieldBox,
  makeLine,
  makePenStroke,
  makeObject,
  makeMediaFrame,
  nonVisualOf,
  makeSmartArt,
  makeChart,
  objectProgIdOf,
  reorderChild,
  setLineSpacingPercent,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  shapeTextOf,
  slideNotesOf,
  toggleBullet,
  toggleNumbering,
  toggleRunFlag,
  toggleRunStyle,
  writeCellText,
  writeShapeText,
  writeSlideNotes,
  THEME_PRESETS,
  variantSchemesOf,
  resetSlidePlaceholders,
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
type Para = {
  properties?: {
    alignment?: string;
    level?: number;
    bullet?: { type: string; char?: string };
    lineSpacingPercent?: number;
  };
  children?: (Run | string)[];
};
const paras = (paragraphs: Para[]): ParagraphDescriptorOptions[] =>
  paragraphs as unknown as ParagraphDescriptorOptions[];

const bodyOf = (paragraphs: unknown[]): TextBodyOptions =>
  ({ paragraphs }) as unknown as TextBodyOptions;

const shapeChild = (body: unknown): SlideChild =>
  ({ shape: { textBody: body } }) as unknown as SlideChild;

const shapeOf = (body: unknown): ShapeVariant => ({ textBody: body }) as unknown as ShapeVariant;

const cellOf = (children: unknown[]): TableCellOptions =>
  ({ children }) as unknown as TableCellOptions;

type ShapeVariant = Extract<SlideChild, { shape: unknown }>["shape"];
type PictureVariant = Extract<SlideChild, { picture: unknown }>["picture"];

describe("toggleBullet / toggleNumbering", () => {
  it("applies the default bullet to every paragraph", () => {
    const paragraphs = paras([{ children: ["a"] }, { children: ["b"] }]);
    toggleBullet(paragraphs);
    expect(paragraphs.map((p) => p.properties?.bullet)).toEqual([
      { type: "char" },
      { type: "char" },
    ]);
  });

  it("clears the bullet when every paragraph carries one", () => {
    const paragraphs = paras([
      { properties: { bullet: { type: "char", char: "•" } }, children: ["a"] },
      { properties: { bullet: { type: "autoNum" } }, children: ["b"] },
    ]);
    toggleBullet(paragraphs);
    expect(paragraphs.map((p) => p.properties?.bullet)).toEqual([undefined, undefined]);
  });

  it("counts buNone as off and applies over it", () => {
    const paragraphs = paras([{ properties: { bullet: { type: "none" } }, children: ["a"] }]);
    toggleNumbering(paragraphs);
    expect(paragraphs[0]!.properties!.bullet).toEqual({ type: "autoNum" });
  });

  it("applies numbering onto a mixed selection", () => {
    const paragraphs = paras([
      { children: ["a"] },
      { properties: { bullet: { type: "char" } }, children: ["b"] },
    ]);
    toggleNumbering(paragraphs);
    expect(paragraphs.map((p) => p.properties?.bullet?.type)).toEqual(["autoNum", "autoNum"]);
  });
});

describe("setLineSpacingPercent", () => {
  it("stamps the percentage over every paragraph", () => {
    const paragraphs = paras([{ children: ["a"] }, { children: ["b"] }]);
    setLineSpacingPercent(paragraphs, 150);
    expect(paragraphs.map((p) => p.properties?.lineSpacingPercent)).toEqual([150, 150]);
  });
});

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
    const shape = shapeOf(bodyOf([{ children: [{ text: "a" }] }, { children: [{ text: "b" }] }]));
    expect(shapeTextOf(shape)).toBe("a\nb");
  });

  it("reads through the text and string sugar", () => {
    expect(shapeTextOf(shapeOf({ text: "hi" }))).toBe("hi");
    expect(shapeTextOf(shapeOf(bodyOf(["one", "two"])))).toBe("one\ntwo");
  });

  it("reads a soft break as a line separator", () => {
    const shape = shapeOf(bodyOf([{ children: [{ text: "a" }, { break: true }, { text: "b" }] }]));
    expect(shapeTextOf(shape)).toBe("a\nb");
  });

  it("returns null for shapes without a text body", () => {
    expect(shapeTextOf({} as ShapeVariant)).toBeNull();
  });

  it("writes one paragraph per line, keeping alignment and first-run style", () => {
    const shape = shapeOf(
      bodyOf([
        { properties: { alignment: "center" }, children: [{ text: "old", bold: true, size: 24 }] },
      ]),
    );
    expect(writeShapeText(shape, "one\ntwo")).toBe(true);
    const body = shape.textBody as TextBodyOptions;
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
    const shape = shapeOf({ text: "before" });
    writeShapeText(shape, "after");
    const body = shape.textBody as TextBodyOptions;
    expect(body.text).toBeUndefined();
    expect(shapeTextOf(shape)).toBe("after");
  });
});

describe("nonVisualOf", () => {
  it("reads cNvPr from every drawing-backed child kind", () => {
    const children = [
      { table: { name: "Table", hidden: true, rows: [] } },
      { chart: { name: "Chart" } },
      { ole: { name: "Workbook", hidden: true } },
    ] as unknown as SlideChild[];
    expect(nonVisualOf(children[0]!)).toMatchObject({ name: "Table", hidden: true });
    expect(nonVisualOf(children[1]!)).toMatchObject({ name: "Chart" });
    expect(nonVisualOf(children[2]!)).toMatchObject({ name: "Workbook", hidden: true });
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

describe("draw-to-place helpers", () => {
  it("lands a text box on the drawn frame", () => {
    const { shape } = makeTextBox(1280, 720, "hello", {
      x: 40,
      y: 60,
      w: 200,
      h: 80,
    }) as { shape: ShapeVariant };
    expect(shape.x).toBe(40 * EMU_PER_PX);
    expect(shape.y).toBe(60 * EMU_PER_PX);
    expect(shape.width).toBe(200 * EMU_PER_PX);
    expect(shape.height).toBe(80 * EMU_PER_PX);
  });

  it("creates a straight line from endpoint direction", () => {
    const { line } = makeLine("line", { x: 30, y: 20, w: 120, h: 70 }) as {
      line: { x1: number; y1: number; x2: number; y2: number };
    };
    expect(line).toMatchObject({
      x1: 30 * EMU_PER_PX,
      y1: 20 * EMU_PER_PX,
      x2: 150 * EMU_PER_PX,
      y2: 90 * EMU_PER_PX,
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

describe("makeTable", () => {
  it("centers a 3×3 grid at 60% width with equal columns", () => {
    type TableVariant = Extract<SlideChild, { table: unknown }>["table"];
    const { table } = makeTable(1280, 720) as { table: TableVariant };
    const widthEmu = Math.round(1280 * 0.6 * EMU_PER_PX);
    expect(table.columnWidths).toEqual([widthEmu / 3, widthEmu / 3, widthEmu / 3]);
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0]!.cells).toHaveLength(3);
    expect(table.x).toBe(Math.round(((1280 - 768) / 2) * EMU_PER_PX));
  });
});

describe("makeObject", () => {
  it("embeds source bytes with a PowerPoint-style icon frame", () => {
    type OleVariant = Extract<SlideChild, { ole: unknown }>["ole"];
    const data = new Uint8Array([1, 2, 3]);
    const { ole } = makeObject(
      1280,
      720,
      data,
      "data:image/png;base64,AAA",
      "Report.xlsx",
      "Excel.Sheet.12",
    ) as { ole: OleVariant };
    expect(ole.x).toBe(Math.round(((1280 - 96) / 2) * EMU_PER_PX));
    expect(ole.y).toBe(Math.round(((720 - 96) / 2) * EMU_PER_PX));
    expect(ole.width).toBe(96 * EMU_PER_PX);
    expect(ole.height).toBe(96 * EMU_PER_PX);
    expect(ole.name).toBe("Report.xlsx");
    expect(ole.progId).toBe("Excel.Sheet.12");
    expect(ole.showAsIcon).toBe(true);
    expect(ole.embed?.data).toBe(data);
    expect(ole.iconImage).toEqual({ data: "data:image/png;base64,AAA", type: "png" });
  });

  it("maps familiar documents to their registered OLE servers", () => {
    expect(objectProgIdOf("Budget.xlsx")).toBe("Excel.Sheet.12");
    expect(objectProgIdOf("Archive.zip")).toBe("CompressedFolder");
    expect(objectProgIdOf("Model.xyz")).toBe("Package");
  });

  it("links a remote source without embedding empty bytes", () => {
    type OleVariant = Extract<SlideChild, { ole: unknown }>["ole"];
    const { ole } = makeObject(
      1280,
      720,
      new Uint8Array(),
      "data:image/png;base64,AAA",
      "Report.xlsx",
      "Excel.Sheet.12",
      { url: "https://example.com/Report.xlsx", autoUpdate: true },
      false,
    ) as { ole: OleVariant };
    expect(ole.link).toEqual({ url: "https://example.com/Report.xlsx", autoUpdate: true });
    expect(ole.embed).toBeUndefined();
    expect(ole.showAsIcon).toBe(false);
  });
});

describe("makeSmartArt", () => {
  it("centers a four-step process diagram with real data-model nodes", () => {
    type SmartArtVariant = Extract<SlideChild, { smartart: unknown }>["smartart"];
    const { smartart } = makeSmartArt(1280, 720) as { smartart: SmartArtVariant };
    expect(smartart.layout).toBe("process1");
    expect(smartart.nodes.map((node) => node.text)).toEqual([
      "Discover",
      "Design",
      "Build",
      "Launch",
    ]);
    expect(smartart.width).toBe(Math.round(768 * EMU_PER_PX));
    expect(smartart.height).toBe(Math.round(720 * 0.48 * EMU_PER_PX));
    expect(smartart.x).toBe(Math.round(256 * EMU_PER_PX));
  });
});

describe("makeChart", () => {
  it("centers a real column-chart model at the shared 60% × 48% frame", () => {
    type ChartVariant = Extract<SlideChild, { chart: unknown }>["chart"];
    const { chart } = makeChart(1280, 720) as { chart: ChartVariant };
    expect(chart.type).toBe("column");
    expect(chart.categories).toHaveLength(4);
    expect(chart.series).toHaveLength(2);
    expect(chart.width).toBe(Math.round(768 * EMU_PER_PX));
    expect(chart.height).toBe(Math.round(720 * 0.48 * EMU_PER_PX));
    expect(chart.x).toBe(Math.round(256 * EMU_PER_PX));
  });
});

describe("theme presets", () => {
  it("ship complete Office palettes with font pairs", () => {
    expect(THEME_PRESETS.length).toBeGreaterThanOrEqual(2);
    for (const preset of THEME_PRESETS) {
      expect(preset.colorScheme.accent1).toMatch(/^[0-9A-F]{6}$/);
      expect(preset.colorScheme.accent6).toMatch(/^[0-9A-F]{6}$/);
      expect(preset.fontScheme?.majorFont?.latin?.typeface).toBeTruthy();
    }
  });

  it("derive four accent rotations from the current scheme", () => {
    const base = THEME_PRESETS[0]!.colorScheme;
    const variants = variantSchemesOf(base);
    expect(variants).toHaveLength(4);
    expect(variants[0]!.accent1).toBe(base.accent1);
    expect(variants[1]!.accent1).toBe(base.accent2);
    expect(variants[3]!.accent1).toBe(base.accent4);
    for (const variant of variants) expect(variant.accent2).toBeTruthy();
  });
});

describe("resetSlidePlaceholders", () => {
  const slideChild = (placeholder: string, geometry: Record<string, number>) =>
    ({ shape: { placeholder, textBody: { text: "keep" }, ...geometry } }) as never;

  it("restores placeholder geometry from the layout, text untouched", () => {
    const slide = {
      children: [
        slideChild("title", { x: 1, y: 2, width: 3, height: 4 }),
        { shape: { x: 9, y: 9, width: 1, height: 1 } },
      ],
    } as never;
    const layout = {
      children: [slideChild("title", { x: 10, y: 20, width: 30, height: 40 })],
    } as never;
    const [moved] = resetSlidePlaceholders(slide, layout, undefined);
    expect(moved).toBeDefined();
    const shape = (
      moved!.child as {
        shape: { x: number; y: number; width: number; textBody?: { text: string } };
      }
    ).shape;
    expect(shape.x).toBe(10);
    expect(shape.y).toBe(20);
    expect(shape.width).toBe(30);
    expect(shape.textBody?.text).toBe("keep");
    expect(moved!.before).toEqual({ x: 1, y: 2, width: 3, height: 4 });
    // The non-placeholder child never moves.
    const children = (slide as { children: { shape: { x: number } }[] }).children;
    expect(children[1]!.shape.x).toBe(9);
  });

  it("falls back to the master when the layout placeholder carries no xfrm", () => {
    const slide = { children: [slideChild("dt", { x: 1, y: 1 })] } as never;
    const layout = { children: [{ shape: { placeholder: "dt" } }] } as never;
    const master = {
      children: [slideChild("dt", { x: 838200, y: 6356350 })],
    } as never;
    const [moved] = resetSlidePlaceholders(slide, layout, master);
    const shape = (moved!.child as { shape: { x: number; y: number } }).shape;
    expect(shape.x).toBe(838200);
    expect(shape.y).toBe(6356350);
  });

  it("returns empty when nothing inherits", () => {
    const slide = { children: [slideChild("title", { x: 1, y: 1 })] } as never;
    expect(resetSlidePlaceholders(slide, undefined, undefined)).toEqual([]);
  });
});

describe("makeMediaFrame", () => {
  it("centers native video bytes in a 16:9 frame", () => {
    const { video } = makeMediaFrame(
      1280,
      720,
      "video",
      new Uint8Array([1]),
      "mp4",
      "clip.mp4",
    ) as {
      video: {
        x: number;
        y: number;
        width: number;
        height: number;
        type: string;
        fileName?: string;
      };
    };
    expect(video.type).toBe("mp4");
    expect(video.fileName).toBe("clip.mp4");
    expect(video.width).toBe(Math.round(640 * EMU_PER_PX));
    expect(video.height).toBe(Math.round(360 * EMU_PER_PX));
    expect(video.x).toBe(Math.round(320 * EMU_PER_PX));
    expect(video.y).toBe(Math.round(180 * EMU_PER_PX));
  });

  it("uses the shorter audio frame", () => {
    const { audio } = makeMediaFrame(1280, 720, "audio", new Uint8Array([1]), "mp3") as {
      audio: { width: number; height: number; type: string };
    };
    expect(audio.type).toBe("mp3");
    expect(audio.width).toBe(Math.round(384 * EMU_PER_PX));
    expect(audio.height).toBe(Math.round(80 * EMU_PER_PX));
  });
});

describe("makeShape", () => {
  it('centers a fresh 2" shape carrying the preset geometry', () => {
    const { shape } = makeShape(1280, 720, "star5") as { shape: ShapeVariant };
    expect(shape.width).toBe(Math.round(192 * EMU_PER_PX));
    expect(shape.x).toBe(Math.round(((1280 - 192) / 2) * EMU_PER_PX));
    expect(shape.properties).toMatchObject({ geometry: "star5", fill: "FFFFFF" });
    expect(shape.textBody).toBeDefined();
  });

  it("uses the drawn rectangle for drag-to-draw insertion", () => {
    const { shape } = makeShape(1280, 720, "star5", {
      x: 50,
      y: 70,
      w: 160,
      h: 90,
    }) as { shape: ShapeVariant };
    expect(shape.x).toBe(50 * EMU_PER_PX);
    expect(shape.y).toBe(70 * EMU_PER_PX);
    expect(shape.width).toBe(160 * EMU_PER_PX);
    expect(shape.height).toBe(90 * EMU_PER_PX);
  });
});

describe("makePenStroke", () => {
  it("writes a literal freeform custGeom path", () => {
    const { shape } = makePenStroke([
      { x: 20, y: 30 },
      { x: 80, y: 90 },
    ]) as { shape: ShapeVariant };
    expect(shape.x).toBe(20 * EMU_PER_PX);
    expect(shape.y).toBe(30 * EMU_PER_PX);
    expect(shape.width).toBe(60 * EMU_PER_PX);
    expect(shape.height).toBe(60 * EMU_PER_PX);
    const geometry = (
      shape.properties as {
        customGeometry: {
          pathList: {
            w: number;
            h: number;
            fill: string;
            stroke: boolean;
            commands: { command: string; point?: { x: string; y: string } }[];
          }[];
        };
      }
    ).customGeometry;
    const path = geometry.pathList[0]!;
    expect(path).toMatchObject({
      w: 60 * EMU_PER_PX,
      h: 60 * EMU_PER_PX,
      fill: "none",
      stroke: true,
    });
    expect(path.commands[0]).toMatchObject({
      command: "moveTo",
      point: { x: "0", y: "0" },
    });
    expect(path.commands[1]).toMatchObject({
      command: "lineTo",
      point: { x: String(60 * EMU_PER_PX), y: String(60 * EMU_PER_PX) },
    });
  });

  it("rejects a tap without a stroke segment", () => {
    expect(makePenStroke([{ x: 1, y: 2 }])).toBeNull();
  });
});

describe("makeFieldBox", () => {
  it("anchors the slide number field bottom-right with its cached text", () => {
    const { shape } = makeFieldBox(1280, 720, "slidenum", "3", "right") as { shape: ShapeVariant };
    expect(shape.x).toBeGreaterThan((1280 / 2) * EMU_PER_PX);
    expect(shape.y).toBeGreaterThan((720 / 2) * EMU_PER_PX);
    expect(shape.properties).toMatchObject({ geometry: "rect", fill: { type: "none" } });
    const paragraph = (shape.textBody as { paragraphs: Para[] }).paragraphs[0]!;
    expect(paragraph.properties?.alignment).toBe("right");
    expect(paragraph.children?.[0]).toEqual({ type: "slidenum", text: "3" });
  });

  it("anchors the date-time field bottom-left", () => {
    const { shape } = makeFieldBox(1280, 720, "datetimeFigureOut", "9/25/2026", "left") as {
      shape: ShapeVariant;
    };
    expect(shape.x).toBeLessThan((1280 / 2) * EMU_PER_PX);
    const paragraph = (shape.textBody as { paragraphs: Para[] }).paragraphs[0]!;
    expect(paragraph.children?.[0]).toMatchObject({ type: "datetimeFigureOut" });
  });
});

describe("slide notes", () => {
  it("reads the string shorthand and the structured text sugar", () => {
    expect(slideNotesOf({ notes: "spoken" })).toBe("spoken");
    expect(slideNotesOf({ notes: { text: "structured" } })).toBe("structured");
    expect(slideNotesOf({})).toBe("");
  });

  it("writes the shorthand, upgrades nothing, and clears", () => {
    const plain: SlideOptions = {};
    writeSlideNotes(plain, "a");
    expect(plain.notes).toBe("a");
    writeSlideNotes(plain, "");
    expect("notes" in plain).toBe(false);

    const structured: SlideOptions = { notes: { text: "old" } };
    writeSlideNotes(structured, "new");
    expect(structured.notes).toEqual({ text: "new" });
    writeSlideNotes(structured, "");
    expect(structured.notes).toEqual({});
  });
});
