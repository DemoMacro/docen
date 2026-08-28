import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCX, type JSONContent } from "../index";

// The ribbon's Text Box / Shapes insert commands build a wpsShape node whose
// geometry (transformation/floating/presetGeometry/fill/outline) rides on
// attrs.wpsShape and whose text body is PM content. This pins the export
// contract end to end: that exact node shape must generate a DOCX that parses
// back into the same geometry with the body restored as content.

/** The text-box variant the Text Box command inserts (defaults per Word: page
 *  centered, wrap none, white fill, accent hairline). */
function textBoxDoc(): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "wpsShape",
            attrs: {
              wpsShape: {
                transformation: { width: 1828800, height: 1097280 },
                floating: {
                  horizontalPosition: { relative: "page", align: "center" },
                  verticalPosition: { relative: "page", align: "center" },
                  wrap: { type: "none" },
                },
                fill: { type: "solid", color: "FFFFFF" },
                outline: { color: "4472C4", width: 12700 },
              },
            },
            content: [{ type: "paragraph" }],
          },
        ],
      },
    ],
  };
}

/** The shape variant the Shapes gallery inserts (a preset with accent fill). */
function shapeDoc(): JSONContent {
  const doc = textBoxDoc();
  const node = doc.content![0].content![0];
  node.attrs!.wpsShape = {
    ...node.attrs!.wpsShape,
    presetGeometry: { preset: "ellipse" },
    fill: { type: "solid", color: "4472C4" },
    outline: { color: "2F528F", width: 12700 },
  };
  return doc;
}

function firstWpsShape(json: JSONContent): {
  attrs: Record<string, unknown>;
  content?: JSONContent[];
} {
  for (const para of json.content ?? []) {
    for (const child of para.content ?? []) {
      if (child.type === "wpsShape") return child as never;
    }
  }
  throw new Error("no wpsShape in parsed document");
}

describe("wpsShape insert round-trip", () => {
  it("restores the text-box geometry from a generated DOCX", () => {
    const json = parseDOCX(generateDOCXSync(textBoxDoc()) as Uint8Array);
    const node = firstWpsShape(json);
    const ws = node.attrs.wpsShape as Record<string, any>;
    // The round-tripped xfrm also carries the effect extents (a/b/l/r/t, all
    // zero) — only the extent itself is insert-command authored.
    expect(ws.transformation).toMatchObject({ width: 1828800, height: 1097280 });
    expect(ws.floating.horizontalPosition).toEqual({ relative: "page", align: "center" });
    expect(ws.floating.verticalPosition).toEqual({ relative: "page", align: "center" });
    expect(ws.floating.wrap).toEqual({ type: "none" });
    expect(ws.fill).toMatchObject({ type: "solid" });
    expect(ws.outline).toMatchObject({ width: 12700 });
    expect(node.content?.length).toBeGreaterThan(0);
  });

  it("restores the preset geometry and accent fill of a gallery shape", () => {
    const json = parseDOCX(generateDOCXSync(shapeDoc()) as Uint8Array);
    const node = firstWpsShape(json);
    const ws = node.attrs.wpsShape as Record<string, any>;
    expect(ws.presetGeometry).toEqual({ preset: "ellipse" });
    expect(ws.fill).toMatchObject({ type: "solid" });
    expect(ws.outline).toMatchObject({ width: 12700 });
  });
});
