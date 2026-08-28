import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCX, type JSONContent } from "../index";

// The ribbon's Link command stamps a link mark on the selection (or inserts
// fresh marked text). This pins the export contract: a marked text run must
// generate a DOCX whose hyperlink container parses back into the editor model
// — text carrying a link mark with the same destination (an https URL verbatim,
// a #name bookmark anchor back as "#name", the mark's title as the tooltip).

function docWithLink(href: string, title?: string): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "portal",
            marks: [
              {
                type: "link",
                attrs: {
                  href,
                  target: href.startsWith("#") ? null : "_blank",
                  ...(title ? { title } : {}),
                },
              },
            ],
          },
        ],
      },
    ],
  };
}

function firstLinkMark(json: JSONContent): Record<string, any> {
  const walk = (node: JSONContent): Record<string, any> | undefined => {
    for (const child of node.content ?? []) {
      if (child.type === "text") {
        const mark = (child.marks ?? []).find((m) => m.type === "link");
        if (mark) return mark.attrs as Record<string, any>;
      }
      const hit = walk(child);
      if (hit) return hit;
    }
    return undefined;
  };
  const hit = walk(json);
  if (!hit) throw new Error("no link mark in parsed document");
  return hit;
}

function firstText(json: JSONContent): string {
  const walk = (node: JSONContent): string | undefined => {
    for (const child of node.content ?? []) {
      if (child.type === "text" && child.text) return child.text;
      const hit = walk(child);
      if (hit) return hit;
    }
    return undefined;
  };
  return walk(json) ?? "";
}

describe("link mark round-trip", () => {
  it("exports an external URL hyperlink and parses back as a marked run", () => {
    const json = parseDOCX(
      generateDOCXSync(docWithLink("https://example.com", "Example")) as Uint8Array,
    );
    const attrs = firstLinkMark(json);
    expect(attrs.href).toBe("https://example.com");
    expect(attrs.title).toBe("Example");
    expect(firstText(json)).toContain("portal");
  });

  it("exports a #bookmark anchor and restores it as an in-page href", () => {
    const json = parseDOCX(generateDOCXSync(docWithLink("#section-1")) as Uint8Array);
    expect(firstLinkMark(json).href).toBe("#section-1");
  });
});
