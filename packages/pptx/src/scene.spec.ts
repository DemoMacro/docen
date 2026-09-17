import type { PresentationOptions } from "@office-open/pptx";
import { describe, expect, it } from "vitest";

import { projectPresentation } from "./scene";

const project = (pres: PresentationOptions) => projectPresentation(pres);

// 1×1 transparent PNG.
const png = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52,
]);
const pngSrc = `data:image/png;base64,${btoa(String.fromCharCode(...png))}`;

describe("slide size", () => {
  it("resolves the named classes at 96 dpi", () => {
    expect(project({ size: "16:9" })).toMatchObject({ widthPx: 1280, heightPx: 720 });
    expect(project({ size: "4:3" })).toMatchObject({ widthPx: 960, heightPx: 720 });
  });

  it("defaults to 16:9 when absent", () => {
    expect(project({})).toMatchObject({ widthPx: 1280, heightPx: 720 });
  });

  it("resolves an explicit EMU size", () => {
    expect(project({ size: { width: 9144000, height: 6858000 } })).toMatchObject({
      widthPx: 960,
      heightPx: 720,
    });
  });
});

describe("shapes", () => {
  it("projects a box preset as a shape member with paint", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              shape: {
                x: 952500,
                y: 952500,
                width: 1905000,
                height: 952500,
                properties: {
                  geometry: "roundRect",
                  fill: { type: "solid", color: "FF0000" },
                  outline: { width: 12700, color: "000000" },
                },
              },
            },
          ],
        },
      ],
    });
    expect(slides[0]!.members).toEqual([
      {
        kind: "shape",
        x: 100,
        y: 100,
        width: 200,
        height: 100,
        preset: "roundRect",
        fill: "FF0000",
        line: { px: expect.closeTo(1.333, 2), color: "000000" },
      },
    ]);
  });

  it("expands a non-box preset through the evaluator into path members", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              shape: {
                x: 0,
                y: 0,
                width: 952500,
                height: 952500,
                properties: { geometry: "heart", fill: "FF0000" },
              },
            },
          ],
        },
      ],
    });
    const members = slides[0]!.members;
    expect(members.length).toBeGreaterThan(0);
    expect(members.every((m) => m.kind === "path")).toBe(true);
    expect(members.every((m) => m.x === 0 && m.width === 100 && m.height === 100)).toBe(true);
    expect(members.some((m) => m.kind === "path" && m.fill === "FF0000")).toBe(true);
  });

  it("projects a text-carrying shape as a text box with silhouette and blocks", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              shape: {
                x: 952500,
                y: 1905000,
                width: 1905000,
                height: 1905000,
                properties: { geometry: "ellipse", fill: "EEEEEE" },
                textBody: {
                  bodyProperties: { anchor: "center" },
                  paragraphs: [
                    {
                      properties: { alignment: "center", spaceBefore: 6 },
                      children: [
                        { text: "Hello", bold: true, size: 24, font: "Arial", fill: "111111" },
                        { break: true },
                        "world",
                      ],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const members = slides[0]!.members;
    expect(members[0]!.kind).toBe("textBox");
    const tb = members[0]! as Extract<(typeof members)[number], { kind: "textBox" }>;
    expect(tb.x).toBe(100);
    expect(tb.y).toBe(200);
    expect(tb.preset).toBe("ellipse");
    expect(tb.d).toBeUndefined(); // a box preset paints natively — no silhouette
    expect(tb.anchor).toBe("center");
    // The DrawingML default insets plus the ellipse's text-rect shrink.
    expect(tb.insets!.left).toBeGreaterThan(9.6);
    const para = tb.blocks[0]! as Extract<(typeof tb.blocks)[number], { kind: "paragraph" }>;
    expect(para.align).toBe("center");
    expect(para.spacing).toEqual({ beforePx: 8, afterPx: 0 });
    expect(para.inline).toHaveLength(3);
    const [run, br, text] = para.inline;
    expect(run).toMatchObject({
      kind: "text",
      text: "Hello",
      style: { bold: true, sizePx: 32, family: "Arial", color: "111111" },
    });
    expect(br).toEqual({ kind: "break" });
    expect(text).toMatchObject({ kind: "text", text: "world" });
  });
});

describe("lines and connectors", () => {
  it("encodes endpoint direction in the diagonal", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              line: {
                x1: 952500,
                y1: 952500,
                x2: 1905000,
                y2: 1905000,
                properties: { outline: { width: 12700, color: "000000" } },
              },
            },
            {
              line: {
                x1: 952500,
                y1: 1905000,
                x2: 1905000,
                y2: 952500,
                properties: { outline: { width: 12700 } },
              },
            },
          ],
        },
      ],
    });
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "path",
      x: 100,
      y: 100,
      d: "M 0 0 L 100 100",
    });
    // The second line rises: its start sits at the box's bottom-left.
    expect(slides[0]!.members[1]).toMatchObject({
      kind: "path",
      x: 100,
      y: 100,
      d: "M 0 100 L 100 0",
    });
  });

  it("expands stroke end arrows into their own members", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              line: {
                x1: 0,
                y1: 0,
                x2: 952500,
                y2: 0,
                properties: {
                  outline: {
                    width: 12700,
                    color: "000000",
                    tailEnd: { type: "triangle" },
                  },
                },
              },
            },
          ],
        },
      ],
    });
    const members = slides[0]!.members;
    expect(members).toHaveLength(2);
    expect(members[1]!.kind).toBe("path");
    // The arrow sits at the line's far end, its own fill member.
    expect(members[1]!.x).toBeGreaterThan(90);
  });
});

describe("groups", () => {
  it("maps children through chOff/chExt scaling and addresses them", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              group: {
                x: 1905000,
                y: 952500,
                width: 1905000,
                height: 952500,
                childOffsetX: 0,
                childOffsetY: 0,
                childExtentWidth: 3810000,
                childExtentHeight: 1905000,
                children: [
                  {
                    shape: {
                      x: 0,
                      y: 0,
                      width: 952500,
                      height: 952500,
                      properties: { geometry: "rect", fill: "00FF00" },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const [m] = slides[0]!.members;
    // The child at (0,0) in a 2×-compressed space lands at the group origin,
    // its extent scaled by 0.5.
    expect(m).toMatchObject({
      kind: "shape",
      x: 200,
      y: 100,
      width: 50,
      height: 50,
      childPath: [0, 0],
    });
  });
});

describe("pictures and background", () => {
  it("projects a picture with src, crop and flips", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                data: png,
                x: 0,
                y: 0,
                width: 952500,
                height: 952500,
                flipHorizontal: true,
                sourceRectangle: { left: 10, right: 20 },
              },
            },
          ],
        },
      ],
    });
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "picture",
      src: pngSrc,
      flipH: true,
      crop: { left: 0.1, top: 0, right: 0.2, bottom: 0 },
    });
  });

  it("carries the slide's solid background", () => {
    const { slides } = project({
      slides: [{ background: { fill: { type: "solid", color: "0B57D0" } } }],
    });
    expect(slides[0]!.background).toBe("0B57D0");
  });
});
