import { encodeBase64 } from "@office-open/core";
import { describe, expect, it } from "vitest";

import {
  compileDocument,
  generateDOCXSync,
  parseDOCXSync,
  resolveDocument,
  type JSONContent,
} from "../index";

// wpgGroup is a content node: the group's own geometry rides attrs.wpgGroup
// (GroupOptions minus children), the members ride PM content through
// group-members (wps → wpsShape, raster/svg → image with a groupXfrm EMU box,
// nested wpg → wpgGroup, chart/contentPart → inlinePassthrough). These pin the
// member round-trip end to end: an authored group node must generate a DOCX
// that parses back into the same member sequence with EMU-exact geometry.

/** An image member's byte source: a PNG signature-prefixed payload. Both the
 *  generate and parse pipelines store media bytes verbatim (no decode), so the
 *  content past the signature is arbitrary — building it in code avoids the
 *  hand-typed-base64-is-corrupt trap. */
const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 1, 2, 3, 4, 5, 6, 7,
]);
const PNG_SRC = `data:image/png;base64,${encodeBase64(PNG_BYTES)}`;

/** A floating group node carrying `members` as PM content. */
function groupDoc(members: JSONContent[]): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "wpgGroup",
            attrs: {
              wpgGroup: {
                transformation: { width: 2419400, height: 1219200 },
                floating: {
                  horizontalPosition: { relative: "page", align: "center" },
                  verticalPosition: { relative: "page", align: "center" },
                  wrap: { type: "none" },
                },
                childOffsetX: 0,
                childOffsetY: 0,
                childExtentWidth: 2419400,
                childExtentHeight: 1219200,
              },
            },
            content: members,
          },
        ],
      },
    ],
  };
}

/** A wps shape member at child-space EMU `left,top` with an editable body. */
function shapeMember(left: number, top: number, text = "member"): JSONContent {
  return {
    type: "wpsShape",
    attrs: {
      wpsShape: {
        transformation: { width: 1143000, height: 571500, offset: { left, top } },
        geometry: "ellipse",
        fill: { type: "solid", color: "4472C4" },
        outline: { color: "2F528F", width: 12700 },
      },
    },
    content: [{ type: "paragraph", content: [{ type: "text", text }] }],
  };
}

/** A picture member whose EMU child-space box rides groupXfrm. */
function imageMember(): JSONContent {
  return {
    type: "image",
    attrs: {
      src: PNG_SRC,
      groupXfrm: { x: 1270000, y: 190500, cx: 952500, cy: 952500 },
      flipH: true,
    },
  };
}

/** A nested group member: attrs.wpgGroup keeps the GroupMediaData child-space
 *  form (MediaDataTransformation) verbatim. */
function nestedGroupMember(): JSONContent {
  return {
    type: "wpgGroup",
    attrs: {
      wpgGroup: {
        type: "wpg",
        transformation: {
          offset: { emus: { x: 1270000, y: 0 }, pixels: { x: 133, y: 0 } },
          emus: { x: 1143000, y: 1143000 },
          pixels: { x: 120, y: 120 },
        },
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 1143000,
        childExtentHeight: 1143000,
      },
    },
    content: [shapeMember(0, 0, "inner")],
  };
}

function firstGroup(json: JSONContent): JSONContent {
  for (const para of json.content ?? []) {
    for (const child of para.content ?? []) {
      if (child.type === "wpgGroup") return child;
    }
  }
  throw new Error("no wpgGroup in parsed document");
}

describe("wpgGroup member round-trip", () => {
  it("restores a shape member with EMU-exact geometry and its text body", () => {
    const json = parseDOCXSync(
      generateDOCXSync(groupDoc([shapeMember(1270000, 635000)])) as Uint8Array,
    );
    const group = firstGroup(json);
    const g = group.attrs!.wpgGroup as Record<string, any>;
    expect(g.transformation).toMatchObject({ width: 2419400, height: 1219200 });
    expect(g.floating.horizontalPosition).toEqual({ relative: "page", align: "center" });
    expect(g.childExtentWidth).toBe(2419400);

    const member = group.content![0];
    expect(member.type).toBe("wpsShape");
    const ws = member.attrs!.wpsShape as Record<string, any>;
    // Child-space offset survives as the shape's transformation.offset, EMU-exact.
    expect(ws.transformation.offset).toEqual({ left: 1270000, top: 635000 });
    expect(ws.transformation).toMatchObject({ width: 1143000, height: 571500 });
    expect(ws.geometry).toEqual({ preset: "ellipse" });
    expect(ws.fill).toMatchObject({ type: "solid" });
    // The editable body round-trips as content.
    expect(JSON.stringify(member.content)).toContain("member");
  });

  // Needs @office-open/docx ≥0.14.5 (office-open da107d4f): earlier emits
  // r:embed="{undefined}" for a group picture child with no fileName, so
  // picture members never parse back.
  describe("picture member round-trip", () => {
    it("restores a picture member's groupXfrm box and flip", () => {
      const json = parseDOCXSync(generateDOCXSync(groupDoc([imageMember()])) as Uint8Array);
      const member = firstGroup(json).content![0];
      expect(member.type).toBe("image");
      const attrs = member.attrs! as Record<string, any>;
      expect(attrs.src.startsWith("data:image/png;base64,")).toBe(true);
      expect(attrs.groupXfrm).toEqual({ x: 1270000, y: 190500, cx: 952500, cy: 952500 });
      expect(attrs.flipH).toBe(true);
    });

    it("keeps a mixed shape+picture member sequence in order", () => {
      const json = parseDOCXSync(
        generateDOCXSync(groupDoc([shapeMember(0, 0), imageMember()])) as Uint8Array,
      );
      const members = firstGroup(json).content!;
      expect(members.map((m) => m.type)).toEqual(["wpsShape", "image"]);
    });
  });

  it("round-trips a nested group with its child-space transformation verbatim", () => {
    const json = parseDOCXSync(generateDOCXSync(groupDoc([nestedGroupMember()])) as Uint8Array);
    const nested = firstGroup(json).content![0];
    expect(nested.type).toBe("wpgGroup");
    const g = nested.attrs!.wpgGroup as Record<string, any>;
    // GroupMediaData form: MediaDataTransformation (offset.emus + emus), EMU-exact.
    expect(g.transformation.offset.emus).toEqual({ x: 1270000, y: 0 });
    expect(g.transformation.emus).toEqual({ x: 1143000, y: 1143000 });
    expect(g.childExtentWidth).toBe(1143000);
    expect(nested.content![0].type).toBe("wpsShape");
    expect(JSON.stringify(nested.content)).toContain("inner");
  });

  it("carries a chart member through inlinePassthrough byte-faithfully (model bridge)", () => {
    // A chart child has no PM node of its own — the member rides an
    // inlinePassthrough atom (same as the paragraph top level). resolve↔compile
    // must restore the exact child object.
    const chartChild = {
      type: "chart",
      transformation: {
        offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
        pixels: { x: 120, y: 90 },
        emus: { x: 1143000, y: 857250 },
      },
      chartKey: "chart-1",
      nonVisualProperties: { id: 2, name: "Chart 2" },
    };
    const docOpts = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    wpgGroup: {
                      transformation: { width: 2419400, height: 1219200 },
                      childOffsetX: 0,
                      childOffsetY: 0,
                      childExtentWidth: 2419400,
                      childExtentHeight: 1219200,
                      children: [chartChild, wpsChildData()],
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const json = resolveDocument(docOpts as never);
    const group = firstGroup(json);
    expect(group.content![0].type).toBe("inlinePassthrough");
    expect(group.content![1].type).toBe("wpsShape");

    const back = compileDocument(json) as typeof docOpts;
    const children = (
      back.sections[0].children[0] as {
        paragraph: { children: { wpgGroup: { children: unknown[] } }[] };
      }
    ).paragraph.children[0].wpgGroup.children;
    expect(children[0]).toEqual(chartChild);
    expect((children[1] as { type: string }).type).toBe("wps");
  });
});

/** A wps group child in office-open form (parse-side shape), for the chart
 *  fixture's sibling — proves a passthrough member coexists with modeled ones. */
function wpsChildData(): {
  type: "wps";
  transformation: {
    offset: { pixels: { x: number; y: number }; emus: { x: number; y: number } };
    pixels: { x: number; y: number };
    emus: { x: number; y: number };
  };
  data: { children: string[]; geometry: string };
} {
  return {
    type: "wps",
    transformation: {
      offset: { pixels: { x: 0, y: 0 }, emus: { x: 1270000, y: 635000 } },
      pixels: { x: 133, y: 66 },
      emus: { x: 1143000, y: 571500 },
    },
    data: { children: ["inner"], geometry: "rect" },
  };
}
