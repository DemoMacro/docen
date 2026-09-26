import { EMU_PER_PX } from "@docen/layout";
import type { PresentationOptions, SlideChild } from "@office-open/pptx";
import { describe, expect, it } from "vitest";

import { projectPresentation } from "./scene";
import { textBlocks } from "./scene/text";
import { memberAt, memberByPath, offsetMemberByPath, resizeMemberByPath } from "./scene/walk";

const project = (pres: PresentationOptions) => projectPresentation(pres);

function groupChildOf(child: SlideChild, index: number) {
  if (!("group" in child)) throw new Error("expected a group child");
  const member = child.group.children?.[index];
  if (!member) throw new Error("expected a group member");
  return member;
}

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

  it("keeps a text-carrying shape's rotation on its whole text box", () => {
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
                rotation: 30,
                properties: { geometry: "rect" },
                textBody: { paragraphs: [{ text: "Spin" }] },
              },
            },
          ],
        },
      ],
    });
    const [member] = slides[0]!.members;
    expect(member).toMatchObject({
      kind: "textBox",
      rotation: 30,
      rotationAbout: "center",
    });
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

  it("mirrors bent connectors when endpoint pairs run reversed", () => {
    const connector = (x1: number, y1: number, x2: number, y2: number) => ({
      connector: {
        x1,
        y1,
        x2,
        y2,
        properties: { geometry: "bentConnector3" as const, outline: { width: 12700 } },
      },
    });
    const { slides } = project({
      slides: [
        {
          children: [
            connector(0, 0, 1905000, 952500),
            connector(1905000, 952500, 0, 0),
            connector(0, 952500, 1905000, 0),
          ],
        },
      ],
    });
    const [forward, reversed, vertical] = slides[0]!.members;
    const pathDataOf = (member: typeof forward | undefined) => {
      if (member?.kind !== "path") throw new Error("expected a path member");
      return member.d;
    };
    expect(forward).toMatchObject({ kind: "path", x: 0, y: 0, width: 200, height: 100 });
    expect("flipH" in forward && forward.flipH).toBe(false);
    expect("flipV" in forward && forward.flipV).toBe(false);
    expect(reversed).toMatchObject({
      kind: "path",
      x: 0,
      y: 0,
      width: 200,
      height: 100,
      flipH: true,
      flipV: true,
    });
    expect(pathDataOf(reversed)).toBe(pathDataOf(forward));
    expect(vertical).toMatchObject({
      kind: "path",
      x: 0,
      y: 0,
      width: 200,
      height: 100,
      flipV: true,
    });
    expect("flipH" in vertical && vertical.flipH).toBe(false);
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

  it("hits members and resolves them by path in slide-absolute px", () => {
    const child: SlideChild = {
      group: {
        x: 2540000,
        y: 1905000,
        width: 5080000,
        height: 3810000,
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 2540000,
        childExtentHeight: 1905000,
        children: [
          {
            shape: {
              x: 0,
              y: 0,
              width: 1270000,
              height: 952500,
              properties: { geometry: "rect", fill: "FF0000" },
            },
          },
          {
            shape: {
              x: 1270000,
              y: 952500,
              width: 1270000,
              height: 952500,
              properties: { geometry: "rect", fill: "00B050" },
            },
          },
        ],
      },
    };
    // The 2× scale maps the red child onto the group's top-left quadrant
    // (266.67,200 → 533.33,400 in slide px).
    const red = memberAt(child, 400, 300);
    expect(red).toMatchObject({
      path: [0],
      x: 2540000 / 9525,
      y: 200,
      width: 2540000 / 9525,
      height: 200,
    });
    expect(red?.scale).toEqual({ sx: 2, sy: 2 });
    // The point between the two members hits nothing (the group itself).
    expect(memberAt(child, 400, 500)).toBeNull();
    const green = memberByPath(child, [1]);
    expect(green).toMatchObject({
      path: [1],
      x: 5080000 / 9525,
      y: 400,
      width: 2540000 / 9525,
      height: 200,
    });
    expect(memberByPath(child, [9])).toBeNull();

    // The dragged slide-space box divides back through the group's 2×
    // affine before landing on the member's own child-space geometry.
    resizeMemberByPath(child, [1], { x: 600, y: 300, width: 300, height: 150 });
    const resized = groupChildOf(child, 1);
    if (!("shape" in resized)) throw new Error("expected a shape member");
    expect(resized.shape).toMatchObject({
      x: (500 / 3) * 9525,
      y: 50 * 9525,
      width: 150 * 9525,
      height: 75 * 9525,
    });
    expect(memberByPath(child, [1])).toMatchObject({
      x: 600,
      y: 300,
      width: 300,
      height: 150,
    });
  });

  it("resizes a line member through the group map without flipping it", () => {
    const child: SlideChild = {
      group: {
        x: 0,
        y: 0,
        width: 5080000,
        height: 3810000,
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 2540000,
        childExtentHeight: 1905000,
        children: [
          {
            line: {
              x1: 0,
              y1: 0,
              x2: 1270000,
              y2: 952500,
            },
          },
        ],
      },
    };
    resizeMemberByPath(child, [0], { x: 100, y: 50, width: 300, height: 150 });
    const line = groupChildOf(child, 0);
    if (!("line" in line)) throw new Error("expected a line member");
    expect(line.line).toMatchObject({
      x1: 50 * 9525,
      y1: 25 * 9525,
      x2: 200 * 9525,
      y2: 100 * 9525,
    });
  });

  it("folds a group's rotation into each flattened member", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              group: {
                x: 0,
                y: 0,
                width: 1905000,
                height: 952500,
                childOffsetX: 0,
                childOffsetY: 0,
                childExtentWidth: 1905000,
                childExtentHeight: 952500,
                rotation: 90,
                children: [
                  {
                    shape: {
                      x: 0,
                      y: 0,
                      width: 952500,
                      height: 952500,
                      properties: { geometry: "rect", fill: "FF0000" },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const [member] = slides[0]!.members;
    expect(member).toMatchObject({
      x: 50,
      y: -50,
      width: 100,
      height: 100,
      rotation: 90,
    });
  });

  it("hits, moves and resizes a member inside a rotated group", () => {
    const child: SlideChild = {
      group: {
        x: 1905000,
        y: 952500,
        width: 1905000,
        height: 952500,
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 1905000,
        childExtentHeight: 952500,
        rotation: 90,
        children: [
          {
            shape: {
              x: 476250,
              y: 238125,
              width: 952500,
              height: 476250,
              properties: { geometry: "rect" },
            },
          },
        ],
      },
    };
    const hit = memberAt(child, 300, 150);
    expect(hit).toMatchObject({
      path: [0],
      x: 250,
      y: 125,
      width: 100,
      height: 50,
      rotation: 90,
      groupRotation: 90,
    });
    expect(memberByPath(child, [0])).toEqual(hit);

    offsetMemberByPath(child, [0], 10, 20);
    const moved = groupChildOf(child, 0);
    if (!("shape" in moved)) throw new Error("expected a shape member");
    expect(moved.shape).toMatchObject({
      x: 70 * 9525,
      y: 15 * 9525,
    });

    resizeMemberByPath(child, [0], { x: 260, y: 145, width: 100, height: 50 });
    const resized = groupChildOf(child, 0);
    if (!("shape" in resized)) throw new Error("expected a shape member");
    expect(resized.shape).toMatchObject({
      x: 70 * 9525,
      y: 15 * 9525,
      width: 100 * 9525,
      height: 50 * 9525,
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

  it("drops cNvPr-hidden children from the projection", () => {
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
                name: "visible",
                properties: { geometry: "rect", fill: "FF0000" },
              },
            },
            {
              shape: {
                x: 952500,
                y: 0,
                width: 952500,
                height: 952500,
                name: "ghost",
                hidden: true,
                properties: { geometry: "rect", fill: "00FF00" },
              },
            },
            {
              table: {
                x: 0,
                y: 1905000,
                width: 1905000,
                height: 952500,
                columnWidths: [952500, 952500],
                rows: [
                  {
                    height: 952500,
                    cells: [{ text: "a" }, { text: "b" }],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    // The visible shape and the table paint; the hidden one drops out.
    expect(slides[0]!.members).toHaveLength(2);
    expect(slides[0]!.members.filter((m) => m.kind === "shape")).toHaveLength(1);
    expect(slides[0]!.members.filter((m) => m.kind === "table")).toHaveLength(1);
  });
});

describe("text bullet projection", () => {
  const bodyOf = (paragraphs: unknown[]) => ({ paragraphs }) as never;

  it("prepends a synthetic bullet glyph with PowerPoint's hanging indent", () => {
    const blocks = textBlocks(
      bodyOf([{ properties: { bullet: { type: "char" } }, children: ["first"] }]),
    );
    const marker = blocks[0]!.inline[0]! as { text: string; synthetic?: boolean };
    const hop = blocks[0]!.inline[1]! as { kind: string };
    expect(marker).toMatchObject({ text: "•", synthetic: true });
    expect(hop).toMatchObject({ kind: "tab" });
    expect(blocks[0]!.indent).toEqual({
      leftPx: 342900 / EMU_PER_PX,
      firstLinePx: -342900 / EMU_PER_PX,
    });
  });

  it("numbers consecutive autoNum paragraphs and restarts after plain text", () => {
    const blocks = textBlocks(
      bodyOf([
        { properties: { bullet: { type: "autoNum" } }, children: ["a"] },
        { properties: { bullet: { type: "autoNum", format: "alphaLcPeriod" } }, children: ["b"] },
        { children: ["plain"] },
        { properties: { bullet: { type: "autoNum" } }, children: ["c"] },
      ]),
    );
    const textOf = (block: (typeof blocks)[number]) => (block.inline[0] as { text?: string }).text;
    expect(textOf(blocks[0]!)).toBe("1.");
    expect(textOf(blocks[1]!)).toBe("a.");
    // The plain paragraph broke the series — the next list restarts at one.
    expect(textOf(blocks[3]!)).toBe("1.");
  });

  it("lands indentLevel on a deeper marL", () => {
    const blocks = textBlocks(
      bodyOf([{ properties: { bullet: { type: "char" }, indentLevel: 1 }, children: ["x"] }]),
    );
    expect(blocks[0]!.indent).toEqual({
      leftPx: (342900 + 457200) / EMU_PER_PX,
      firstLinePx: -(342900 + 457200) / EMU_PER_PX,
    });
  });

  it("projects line spacing as an exact or multiple line height", () => {
    const blocks = textBlocks(
      bodyOf([
        { properties: { lineSpacingPercent: 150 }, children: ["a"] },
        { properties: { lineSpacingPoints: 24 }, children: ["b"] },
      ]),
    );
    expect(blocks[0]!.spacing?.lineHeight).toEqual({ rule: "multiple", factor: 1.5 });
    expect(blocks[1]!.spacing?.lineHeight).toEqual({ rule: "exact", px: 32 });
  });
});

describe("media frame projection", () => {
  it("wraps the poster in a browser URL and preserves the media family", () => {
    const video: SlideChild = {
      video: {
        x: 0,
        y: 0,
        width: 1905000,
        height: 952500,
        data: new Uint8Array([1, 2, 3]),
        type: "mp4",
        poster: new Uint8Array([1, 2, 3]),
        posterType: "png",
        fileName: "clip.mp4",
      },
    } as unknown as SlideChild;
    const { slides } = project({ slides: [{ children: [video] }] });
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "mediaFrame",
      media: "video",
      width: 200,
      height: 100,
      src: "data:image/png;base64,AQID",
      fileName: "clip.mp4",
    });
  });

  it("gives audio without a poster the stable dark-player payload", () => {
    const audio: SlideChild = {
      audio: {
        x: 0,
        y: 0,
        width: 1905000,
        height: 476250,
        type: "mp3",
        fileName: "voice.mp3",
      },
    } as unknown as SlideChild;
    const { slides } = project({ slides: [{ children: [audio] }] });
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "mediaFrame",
      media: "audio",
      height: 50,
      fileName: "voice.mp3",
    });
    expect(slides[0]!.members[0]).not.toHaveProperty("src");
  });
});

describe("smartart fallback projection", () => {
  const smartart: SlideChild = {
    smartart: {
      x: 0,
      y: 0,
      width: 4762500,
      height: 1905000,
      layout: "process1",
      nodes: [{ text: "Start" }, { text: "Middle" }, { text: "End" }],
    },
  } as unknown as SlideChild;

  it("projects the diagram family and node tree", () => {
    const { slides } = project({ slides: [{ children: [smartart] }] });
    expect(slides[0]!.members).toHaveLength(1);
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "smartArt",
      width: 500,
      height: 200,
      layout: "process1",
      nodes: [{ text: "Start" }, { text: "Middle" }, { text: "End" }],
    });
  });

  it("addresses nested diagram members through their group path", () => {
    const child: SlideChild = {
      group: {
        x: 0,
        y: 0,
        width: 9525000,
        height: 1905000,
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 4762500,
        childExtentHeight: 1905000,
        children: [smartart],
      },
    };
    const hit = memberAt(child, 600, 100);
    expect(hit).toMatchObject({ path: [0], x: 0, y: 0, width: 1000, height: 200 });
  });
});
describe("live field projection", () => {
  const fieldBody = (type: string, cached: string) =>
    ({ paragraphs: [{ children: [{ type, text: cached }] }] }) as never;

  it("evaluates slide numbers per projected slide", () => {
    const shape = (number: string): SlideChild =>
      ({
        shape: {
          x: 0,
          y: 0,
          width: 952500,
          height: 952500,
          textBody: {
            paragraphs: [{ children: [{ type: "slidenum", text: number }] }],
          },
        },
      }) as never;
    const { slides } = project({
      slides: [{ children: [shape("7")] }, { children: [shape("7")] }],
    });
    const textOf = (index: number) => {
      const box = slides[index]!.members[0]!;
      if (box.kind !== "textBox") throw new Error("expected a text box");
      const paragraph = box.blocks[0];
      if (paragraph?.kind !== "paragraph") throw new Error("expected a paragraph");
      return paragraph.inline[0]!;
    };
    expect(textOf(0)).toMatchObject({ text: "1" });
    expect(textOf(1)).toMatchObject({ text: "2" });
  });

  it("formats the current date and keeps unknown field caches", () => {
    const date = textBlocks(fieldBody("datetimeFigureOut", "cached"), {
      now: new Date(2026, 8, 27),
    })[0]!.inline[0]!;
    const unknown = textBlocks(fieldBody("custom", "cached"))[0]!.inline[0]!;
    expect(date).toMatchObject({ text: "9/27/2026" });
    expect(unknown).toMatchObject({ text: "cached" });
  });
});

describe("table style projection", () => {
  const tableChild = (table: Record<string, unknown>): SlideChild => ({ table }) as never;
  const tableOf = (pres: PresentationOptions) => {
    const { slides } = project(pres);
    const member = slides[0]!.members[0]!;
    if (member.kind !== "table") throw new Error("expected a table member");
    return member.table as {
      rows: {
        cells: {
          fill?: string;
          borders?: Record<string, unknown>;
          blocks: { kind: string; inline?: { kind: string; style?: unknown }[] }[];
        }[];
      }[];
    };
  };
  const inlineStyleOf = (cell: {
    blocks: { kind: string; inline?: { kind: string; style?: unknown }[] }[];
  }) => {
    const block = cell.blocks[0]!;
    if (block.kind !== "paragraph" || !block.inline) throw new Error("expected a paragraph");
    return block.inline[0]!.style;
  };
  const grid = {
    x: 0,
    y: 0,
    width: 1905000,
    height: 2857500,
    columnWidths: [952500, 952500],
    rows: [
      { height: 952500, cells: [{ text: "h1" }, { text: "h2" }] },
      { height: 952500, cells: [{ text: "a" }, { text: "b" }] },
      { height: 952500, cells: [{ text: "c" }, { text: "d" }] },
    ],
  };

  it("applies the themed default family to a flagged table", () => {
    const table = tableOf({
      masters: [{ theme: { colorScheme: { accent1: "FF0000" } } }] as never,
      slides: [{ children: [tableChild({ ...grid, firstRow: true, bandRow: true })] }],
    });
    const [header, band1, band2] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ fill: "FF0000" });
    expect(inlineStyleOf(header![0]!)).toMatchObject({ bold: true, color: "FFFFFF" });
    // Data rows alternate the light-accent bands; inside rules are white.
    expect(band1![0]).toMatchObject({ fill: "FFCCCC", borders: { top: { color: "FFFFFF" } } });
    expect(band2![0]).toMatchObject({ fill: "FF9999" });
  });

  it("keeps the no-style GUID bare but honors explicit cell fills", () => {
    const table = tableOf({
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              tableStyleId: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
              firstRow: true,
              bandRow: true,
              rows: [
                {
                  height: 952500,
                  cells: [{ text: "h", fill: { type: "solid", color: "00FF00" } }, { text: "h2" }],
                },
                { height: 952500, cells: [{ text: "a" }, { text: "b" }] },
              ],
            }),
          ],
        },
      ],
    });
    const [header, body] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ fill: "00FF00" });
    expect(body![0]!.fill).toBeUndefined();
    expect(body![0]!.borders).toBeUndefined();
  });

  it("resolves a tblStyleLst entry by GUID and layers the flags", () => {
    const guid = "{11111111-2222-3333-4444-555555555555}";
    const table = tableOf({
      tableStyles: {
        defaultStyleId: guid,
        styles: [
          {
            styleId: guid,
            styleName: "Custom",
            regions: {
              wholeTbl: { cell: { fill: '<a:srgbClr val="EEEEEE"/>' } },
              firstRow: { cell: { fill: '<a:srgbClr val="222222"/>' } },
            },
          },
        ],
      },
      slides: [{ children: [tableChild({ ...grid, tableStyleId: guid, firstRow: true })] }],
    });
    const [header, body] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ fill: "222222" });
    expect(body![0]).toMatchObject({ fill: "EEEEEE" });
  });
});
