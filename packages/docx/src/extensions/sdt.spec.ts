import type { DocumentOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";

/**
 * Content-control (w:sdt) container round-trips: the control's settings ride
 * attrs verbatim on both legs, the content resolves to editable nodes and
 * compiles back through the shared block/inline walks.
 */

function roundTrip(children: SectionChild[]) {
  const doc: DocumentOptions = { sections: [{ children }] };
  const json = resolveDocument(doc, docxExtensions);
  const compiled = compileDocument(json, docxExtensions);
  return { json, child: compiled.sections[0].children[0] };
}

describe("sdtBlock", () => {
  it("carries settings verbatim and compiles content back", () => {
    const properties = {
      id: 1,
      tag: "status",
      alias: "Status",
      checkbox: { checked: true },
    };
    const { json, child } = roundTrip([
      { sdt: { properties, children: [{ paragraph: { text: "done" } }] } },
    ]);
    const node = json.content?.[0] as { type: string; attrs?: { properties?: unknown } };
    expect(node.type).toBe("sdtBlock");
    expect(node.attrs?.properties).toEqual(properties);
    expect(child).toEqual({
      sdt: { properties, children: [{ paragraph: "done" }] },
    });
  });

  it("resolves every content paragraph and nests inline controls", () => {
    const { json, child } = roundTrip([
      {
        sdt: {
          properties: { tag: "letter" },
          children: [
            { paragraph: { text: "first" } },
            {
              paragraph: {
                children: [{ sdt: { properties: { alias: "pick" }, children: ["inner"] } }],
              },
            },
          ],
        },
      },
    ]);
    const node = json.content?.[0] as { content?: { type: string }[] };
    expect(node.content?.map((c) => c.type)).toEqual(["paragraph", "paragraph"]);
    const out = (child as { sdt: { children: unknown[] } }).sdt.children;
    expect(out).toHaveLength(2);
    const nested = out[1] as { paragraph: { children: { sdt: unknown }[] } };
    expect(nested.paragraph.children[0].sdt).toEqual({
      properties: { alias: "pick" },
      children: [{ text: "inner" }],
    });
  });

  it("keeps an empty control resolvable with a placeholder paragraph", () => {
    const { json } = roundTrip([{ sdt: { properties: { tag: "empty" } } }]);
    const node = json.content?.[0] as { type: string; content?: unknown[] };
    expect(node.type).toBe("sdtBlock");
    expect(node.content?.length).toBeGreaterThan(0);
  });

  it("round-trips the end-mark run properties", () => {
    const { child } = roundTrip([
      {
        sdt: {
          properties: { tag: "t" },
          children: [{ paragraph: { text: "x" } }],
          endProperties: { bold: true },
        },
      },
    ]);
    expect(child).toMatchObject({ sdt: { endProperties: { bold: true } } });
  });
});

describe("sdtInline", () => {
  it("carries settings verbatim and compiles the run stream back", () => {
    const { json, child } = roundTrip([
      {
        paragraph: {
          children: ["before ", { sdt: { properties: { alias: "pick" }, children: ["body"] } }],
        },
      },
    ]);
    const para = json.content?.[0] as { content?: { type: string }[] };
    expect(para.content?.map((c) => c.type)).toEqual(["text", "sdtInline"]);
    const out = (child as { paragraph: { children: unknown[] } }).paragraph.children;
    expect(out[1]).toEqual({
      sdt: { properties: { alias: "pick" }, children: [{ text: "body" }] },
    });
  });
});
