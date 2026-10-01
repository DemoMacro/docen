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
        sourceChildIndex: 0,
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

describe("custom geometry", () => {
  it("projects literal custGeom paths as scalable outlines", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            {
              shape: {
                x: 0,
                y: 0,
                width: 10 * EMU_PER_PX,
                height: 5 * EMU_PER_PX,
                properties: {
                  customGeometry: {
                    pathList: [
                      {
                        w: 10 * EMU_PER_PX,
                        h: 5 * EMU_PER_PX,
                        fill: "none",
                        stroke: true,
                        commands: [
                          { command: "moveTo", point: { x: "0", y: "0" } },
                          {
                            command: "lineTo",
                            point: { x: String(10 * EMU_PER_PX), y: String(5 * EMU_PER_PX) },
                          },
                        ],
                      },
                    ],
                  },
                  fill: { type: "none" },
                  outline: { width: 12700, color: "262626" },
                },
              },
            },
          ],
        },
      ],
    });
    expect(slides[0]?.members).toEqual([
      expect.objectContaining({
        kind: "path",
        x: 0,
        y: 0,
        width: 10,
        height: 5,
        d: expect.stringContaining("L 10.000 5.000"),
        line: expect.objectContaining({ color: "262626", px: 4 / 3 }),
      }),
    ]);
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

describe("shape fills", () => {
  it("projects a gradient fill as renderer-native stops", () => {
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
                properties: {
                  geometry: "rect",
                  fill: {
                    type: "gradient",
                    angle: 90,
                    stops: [
                      { position: 0, color: "FF0000" },
                      { position: 100, color: "FFFFFF" },
                    ],
                  },
                },
              },
            },
          ],
        },
      ],
    } as never);
    expect(slides[0]!.members[0]).toMatchObject({
      kind: "shape",
      fill: {
        type: "linear",
        from: { x: 50, y: 0 },
        to: { x: 50, y: 100 },
        stops: [
          { offset: 0, color: "#FF0000" },
          { offset: 1, color: "#FFFFFF" },
        ],
      },
    });
  });

  it("projects a theme-colored pattern fill as a tiled image", () => {
    const { slides } = project({
      masters: [{ theme: { colorScheme: { accent1: "FF0000" } } }] as never,
      slides: [
        {
          children: [
            {
              shape: {
                x: 0,
                y: 0,
                width: 952500,
                height: 952500,
                properties: {
                  geometry: "rect",
                  fill: {
                    type: "pattern",
                    pattern: "cross",
                    foregroundColor: { value: "accent1" },
                    backgroundColor: "FFFFFF",
                  },
                },
              },
            },
          ],
        },
      ],
    } as never);
    const member = slides[0]!.members[0]!;
    if (member.kind !== "shape") throw new Error("expected a shape");
    const fill = member.fill;
    if (typeof fill !== "object" || fill.type !== "image")
      throw new Error("expected a pattern paint");
    expect(fill).toMatchObject({ type: "image", mode: "repeat", repeat: true });
    const svg = decodeURIComponent(fill.url.split(",")[1]!);
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain('fill="#FFFFFF"');
    expect(svg).toContain('d="M0 .5H8 M.5 0V8"');
  });

  it("carries color alpha through shape and background paints", () => {
    const { slides } = project({
      masters: [{ theme: { colorScheme: { accent1: "FF0000" } } }] as never,
      slides: [
        {
          background: {
            fill: { type: "solid", color: { value: "FF0000", transforms: { alpha: 25 } } },
          },
          children: [
            {
              shape: {
                x: 0,
                y: 0,
                width: 952500,
                height: 952500,
                properties: {
                  geometry: "rect",
                  fill: {
                    type: "pattern",
                    pattern: "cross",
                    foregroundColor: { value: "accent1", transforms: { alpha: 50 } },
                    backgroundColor: "FFFFFF",
                  },
                },
              },
            },
          ],
        },
        {
          background: {
            fill: {
              type: "gradient",
              angle: 90,
              stops: [
                { position: 0, color: { value: "FF0000", transforms: { alpha: 75 } } },
                { position: 1, color: "FFFFFF" },
              ],
            },
          },
        },
      ],
    } as never);
    expect(slides[0]!.background).toEqual({
      kind: "solid",
      color: "rgba(255, 0, 0, 0.25)",
    });
    const member = slides[0]!.members[0]!;
    if (member.kind !== "shape" || typeof member.fill !== "object" || member.fill.type !== "image")
      throw new Error("expected a pattern paint");
    const svg = decodeURIComponent(member.fill.url.split(",")[1]!);
    expect(svg).toContain('fill="rgba(255, 0, 0, 0.5)"');
    expect(slides[1]!.background).toMatchObject({
      kind: "gradient",
      stops: [
        { color: "rgba(255, 0, 0, 0.75)", position: 0 },
        { color: "FFFFFF", position: 1 },
      ],
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

  it("projects blip tile transforms for shapes and backgrounds", () => {
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
              },
            },
            {
              shape: {
                x: 0,
                y: 0,
                width: 952500,
                height: 952500,
                properties: {
                  geometry: "rect",
                  fill: {
                    type: "blip",
                    data: png,
                    imageType: "png",
                    tile: {
                      tx: 9525,
                      ty: 19050,
                      sx: 50,
                      sy: 150,
                      alignment: "bottomRight",
                    },
                  },
                },
              },
            },
          ],
        },
        {
          background: {
            fill: {
              type: "blip",
              data: png,
              imageType: "png",
              tile: { sx: 25, alignment: "topRight" },
            },
          },
        },
      ],
    } as never);
    expect(slides[0]!.members[1]).toMatchObject({
      kind: "shape",
      fill: {
        type: "image",
        mode: "repeat",
        repeat: true,
        scale: { x: 0.5, y: 1.5 },
        offset: { x: 1, y: 2 },
        align: "bottom-right",
      },
    });
    expect(slides[1]!.background).toMatchObject({
      kind: "image",
      tile: { scale: { x: 0.25, y: 1 }, align: "top-right" },
    });
  });

  it("carries the slide's solid background", () => {
    const { slides } = project({
      slides: [{ background: { fill: { type: "solid", color: "0B57D0" } } }],
    });
    expect(slides[0]!.background).toEqual({ kind: "solid", color: "0B57D0" });
  });

  it("projects a linear gradient's stops and angle", () => {
    const { slides } = project({
      slides: [
        {
          background: {
            fill: {
              type: "gradient",
              angle: 90,
              stops: [
                { position: 0, color: "0B57D0" },
                { position: 1, color: "FFFFFF" },
              ],
            },
          },
        },
      ],
    });
    expect(slides[0]!.background).toEqual({
      kind: "gradient",
      angle: 90,
      stops: [
        { color: "0B57D0", position: 0 },
        { color: "FFFFFF", position: 1 },
      ],
    });
  });

  it("projects the gradient options form's path shade", () => {
    const { slides } = project({
      slides: [
        {
          background: {
            fill: {
              type: "gradient",
              options: {
                stops: [
                  { position: 0, color: { value: "112233" } },
                  { position: 1, color: { value: "FFFFFF" } },
                ],
                shade: { path: "circle" },
              },
            },
          },
        },
      ],
    });
    expect(slides[0]!.background).toMatchObject({ kind: "gradient", path: "circle" });
  });

  it("projects a picture background to a data URL", () => {
    const { slides } = project({
      slides: [{ background: { fill: { type: "blip", data: png, imageType: "png" } } }],
    } as never);
    expect(slides[0]!.background).toEqual({ kind: "image", src: pngSrc });
  });

  it("tiles a pattern fill's preset geometry and colors", () => {
    const { slides } = project({
      slides: [
        {
          background: {
            fill: {
              type: "pattern",
              pattern: "cross",
              foregroundColor: "FF0000",
              backgroundColor: "00FF00",
            },
          },
        },
      ],
    } as never);
    const background = slides[0]!.background;
    if (background?.kind !== "image") throw new Error("expected a pattern image");
    const svg = decodeURIComponent(background.src.split(",")[1]!);
    expect(svg).toContain('pattern id="pattern"');
    expect(svg).toContain('patternUnits="userSpaceOnUse"');
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain('fill="#00FF00"');
    expect(svg).toContain('d="M0 .5H8 M.5 0V8"');
  });

  it("resolves phClr and theme colors in a bgRef pattern", () => {
    const { slides } = project({
      masters: [
        {
          theme: {
            colorScheme: { dark1: "111111" },
            formatScheme: {
              backgroundFillStyles: [
                {
                  type: "pattern",
                  pattern: "diagonalCross",
                  foregroundColor: { value: "phClr" },
                  backgroundColor: { value: "tx1" },
                },
              ],
              fillStyles: [],
              lineStyles: [],
              effectStyles: [],
            },
          },
        },
      ] as never,
      slides: [{ background: { reference: { index: 1001, color: "AA0000" } } }],
    } as never);
    const background = slides[0]!.background;
    if (background?.kind !== "image") throw new Error("expected a pattern image");
    const svg = decodeURIComponent(background.src.split(",")[1]!);
    expect(svg).toContain('fill="#AA0000"');
    expect(svg).toContain('fill="#111111"');
  });

  it("resolves a bgRef through the theme's background fill styles", () => {
    const { slides } = project({
      masters: [
        {
          theme: {
            colorScheme: { light1: "F8F9FA", accent1: "0B57D0" },
            formatScheme: {
              backgroundFillStyles: [
                { type: "solid", color: { value: "phClr" } },
                {
                  type: "gradient",
                  shade: { angle: 90 },
                  stops: [
                    { position: 0, color: { value: "phClr" } },
                    { position: 1, color: { value: "accent1" } },
                  ],
                },
              ],
              fillStyles: [],
              lineStyles: [],
              effectStyles: [],
            },
          },
        },
      ] as never,
      slides: [{ background: { reference: { index: 1001, color: { value: "bg1" } } } }],
    });
    // 1001 → the first bg style; its phClr takes the reference's color, and
    // bg1 maps (through the default color map) to the theme's light1.
    expect(slides[0]!.background).toEqual({ kind: "solid", color: "F8F9FA" });
  });

  it("resolves the second bg style's gradient with phClr stops", () => {
    const { slides } = project({
      masters: [
        {
          theme: {
            colorScheme: { accent1: "112233", accent2: "EEDDCC" },
            formatScheme: {
              backgroundFillStyles: [
                { type: "solid", color: { value: "phClr" } },
                {
                  type: "gradient",
                  shade: { angle: 90 },
                  stops: [
                    { position: 0, color: { value: "phClr" } },
                    { position: 1, color: { value: "accent2" } },
                  ],
                },
              ],
              fillStyles: [],
              lineStyles: [],
              effectStyles: [],
            },
          },
        },
      ] as never,
      slides: [{ background: { reference: { index: 1002, color: { value: "accent1" } } } }],
    });
    expect(slides[0]!.background).toEqual({
      kind: "gradient",
      angle: 90,
      stops: [
        { color: "112233", position: 0 },
        { color: "EEDDCC", position: 1 },
      ],
    });
  });

  it("evaluates phClr transforms in a bgRef fill", () => {
    const { slides } = project({
      masters: [
        {
          theme: {
            colorScheme: { accent1: "FF0000" },
            formatScheme: {
              backgroundFillStyles: [
                {
                  type: "solid",
                  color: { value: "phClr", transforms: { lumMod: 20, lumOff: 80 } },
                },
              ],
              fillStyles: [],
              lineStyles: [],
              effectStyles: [],
            },
          },
        },
      ] as never,
      slides: [{ background: { reference: { index: 1001, color: { value: "accent1" } } } }],
    });
    expect(slides[0]!.background).toEqual({ kind: "solid", color: "FFCCCC" });
  });

  it("inherits the master's bgRef when the slide declares no background", () => {
    const { slides } = project({
      masters: [
        {
          background: { reference: { index: 1001, color: { value: "tx1" } } },
          theme: {
            colorScheme: { dark1: "1A1A1A" },
            formatScheme: {
              backgroundFillStyles: [{ type: "solid", color: { value: "phClr" } }],
              fillStyles: [],
              lineStyles: [],
              effectStyles: [],
            },
          },
        },
      ] as never,
      slides: [{}],
    });
    expect(slides[0]!.background).toEqual({ kind: "solid", color: "1A1A1A" });
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

  it("carries the playable data URL for the browser-decodable formats", () => {
    const video: SlideChild = {
      video: {
        x: 0,
        y: 0,
        width: 1905000,
        height: 952500,
        data: new Uint8Array([1, 2, 3]),
        type: "mp4",
      },
    } as unknown as SlideChild;
    const { slides } = project({ slides: [{ children: [video] }] });
    expect(slides[0]!.members[0]).toMatchObject({
      playable: { src: "data:video/mp4;base64,AQID", mime: "video/mp4" },
    });
  });

  it("keeps the exotic containers unplayable", () => {
    const video: SlideChild = {
      video: {
        x: 0,
        y: 0,
        width: 1905000,
        height: 952500,
        data: new Uint8Array([1, 2, 3]),
        type: "wmv",
      },
    } as unknown as SlideChild;
    const { slides } = project({ slides: [{ children: [video] }] });
    expect(slides[0]!.members[0]).not.toHaveProperty("playable");
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

  it("leaves zero-height auto rows for the painter to grow", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              rows: [
                { height: 0, cells: [{ text: "a" }, { text: "b" }] },
                { height: 0, cells: [{ text: "c" }, { text: "d" }] },
              ],
            }),
          ],
        },
      ],
    });
    const member = slides[0]!.members[0]!;
    if (member.kind !== "table") throw new Error("expected a table member");
    const table = member.table as { rows: { heightPx: number }[] };
    // Projection has no browser canvas in Node; the painter repeats this
    // layout with FontMetrics and grows both rows from cell content.
    expect(table.rows.map((row) => row.heightPx)).toEqual([0, 0]);
  });

  it("splits the frame over the walked grid when declared columns mismatch", () => {
    const { slides } = project({
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              columnWidths: [952500],
              rows: [
                {
                  height: 952500,
                  cells: [{ text: "wide", columnSpan: 2 }, { text: "right" }],
                },
              ],
            }),
          ],
        },
      ],
    });
    const member = slides[0]!.members[0]!;
    if (member.kind !== "table") throw new Error("expected a table member");
    const table = member.table as { columnWidthsPx: number[] };
    expect(table.columnWidthsPx).toHaveLength(3);
    expect(table.columnWidthsPx[0]).toBeCloseTo(200 / 3, 7);
    expect(table.columnWidthsPx[1]).toBeCloseTo(200 / 3, 7);
    expect(table.columnWidthsPx[2]).toBeCloseTo(200 / 3, 7);
  });

  it("renders an unstyled table as PowerPoint's plain grid", () => {
    const table = tableOf({
      masters: [{ theme: { colorScheme: { accent1: "FF0000" } } }] as never,
      slides: [{ children: [tableChild({ ...grid, firstRow: true, bandRow: true })] }],
    });
    const [header, band1, band2] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ borders: { top: { color: "000000" } } });
    expect(header![0]!.fill).toBeUndefined();
    expect(inlineStyleOf(header![0]!)).toEqual({ family: "Calibri", sizePx: 24 });
    expect(band1![0]!.fill).toBeUndefined();
    expect(band2![0]!.fill).toBeUndefined();
  });

  it("applies a derived Medium Style 2 GUID to its replacement accent", () => {
    const table = tableOf({
      masters: [{ theme: { colorScheme: { accent1: "FF0000", accent2: "00FF00" } } }] as never,
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              tableStyleId: "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
              firstRow: true,
              bandRow: true,
            }),
          ],
        },
      ],
    });
    const [header, band1, band2] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ fill: "00FF00" });
    expect(band1![0]).toMatchObject({ fill: "66FF66" });
    expect(band2![0]).toMatchObject({ fill: "33FF33" });
  });

  it("resolves Light Style 1 and Dark Style 1 accent variants", () => {
    const theme = { colorScheme: { accent2: "00AA00", accent3: "0000AA" } };
    const light = tableOf({
      masters: [{ theme }] as never,
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              tableStyleId: "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
              firstRow: true,
            }),
          ],
        },
      ],
    });
    expect(light.rows[0]!.cells[0]!.fill).toBeUndefined();
    expect(light.rows[0]!.cells[0]!.borders).toMatchObject({ top: { color: "00AA00" } });

    const dark = tableOf({
      masters: [{ theme }] as never,
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              tableStyleId: "{D03447BB-5D67-496B-8E87-E561075AD55C}",
              firstRow: true,
            }),
          ],
        },
      ],
    });
    expect(dark.rows[0]!.cells[0]).toMatchObject({ fill: "0000AA" });
    expect(dark.rows[1]!.cells[0]).toMatchObject({ fill: "3333BB" });
    expect(inlineStyleOf(dark.rows[0]!.cells[0]!)).toMatchObject({ bold: true, color: "FFFFFF" });
  });

  it("keeps Dark Style 2's two-accent derivation split by region", () => {
    const table = tableOf({
      masters: [{ theme: { colorScheme: { accent1: "FF0000", accent2: "00FF00" } } }] as never,
      slides: [
        {
          children: [
            tableChild({
              ...grid,
              tableStyleId: "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}",
              firstRow: true,
              bandRow: true,
            }),
          ],
        },
      ],
    });
    const [header, band1, band2] = table.rows.map((row) => row.cells);
    expect(header![0]).toMatchObject({ fill: "00FF00" });
    expect(band1![0]).toMatchObject({ fill: "FF6666" });
    expect(band2![0]).toMatchObject({ fill: "FF3333" });
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
