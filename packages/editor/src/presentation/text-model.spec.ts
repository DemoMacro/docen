import type { SlideChild } from "@docen/pptx";
import type { ParagraphDescriptorOptions } from "@office-open/core/drawing";
import { describe, expect, it } from "vitest";

import { formatTextModel, textModelOf, writeTextModel } from "./text-model";

type AnyShape = Extract<SlideChild, { shape: unknown }>["shape"];

const shapeSource = (paragraphs: unknown[]) =>
  ({
    kind: "shape",
    shape: { textBody: { paragraphs } },
  }) as unknown as { kind: "shape"; shape: AnyShape };

describe("text model", () => {
  it("preserves untouched styled runs while editing plain text", () => {
    const source = shapeSource([{ children: [{ text: "plain " }, { text: "bold", bold: true }] }]);
    writeTextModel(source, "in bold", "plain bold");
    expect(textModelOf(source)).toBe("in bold");
    const paragraphs = source.shape.textBody?.paragraphs as
      | ParagraphDescriptorOptions[]
      | undefined;
    expect(paragraphs?.[0]).toMatchObject({
      children: [{ text: "in " }, { text: "bold", bold: true }],
    });
  });

  it("splits styled runs at the formatted range", () => {
    const source = shapeSource([{ children: [{ text: "plain bold", bold: true }] }]);
    expect(formatTextModel(source, 0, 6, "bold")).toBe(true);
    const paragraphs = source.shape.textBody?.paragraphs as ParagraphDescriptorOptions[];
    expect(paragraphs[0]?.children).toEqual([{ text: "plain " }, { text: "bold", bold: true }]);
  });

  it("keeps zero-width field anchors through a text rewrite", () => {
    const field = { type: "slidenum", text: "1" };
    const source = shapeSource([{ children: [field, { text: "bold", bold: true }] }]);
    writeTextModel(source, "after", "bold");
    const paragraphs = source.shape.textBody?.paragraphs as ParagraphDescriptorOptions[];
    expect(paragraphs[0]?.children?.[0]).toEqual(field);
    expect(paragraphs[0]?.children?.[1]).toMatchObject({ text: "after", bold: true });
  });

  it("keeps paragraph properties when Enter creates a new paragraph", () => {
    const properties = { alignment: "center", bullet: { type: "char" } };
    const source = shapeSource([{ properties, children: [{ text: "before" }] }]);
    writeTextModel(source, "before\nafter", "before");
    const paragraphs = source.shape.textBody?.paragraphs as ParagraphDescriptorOptions[];
    expect(textModelOf(source)).toBe("before\nafter");
    expect(paragraphs?.[0]?.properties).toMatchObject(properties);
    expect(paragraphs?.[1]?.properties).toMatchObject(properties);
  });
});
