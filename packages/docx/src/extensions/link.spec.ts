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

/** The first text node carrying a link mark (marks included, for style checks). */
function firstRunWithLink(json: JSONContent): JSONContent {
  const walk = (node: JSONContent): JSONContent | undefined => {
    for (const child of node.content ?? []) {
      if (child.type === "text" && (child.marks ?? []).some((m) => m.type === "link")) return child;
      const hit = walk(child);
      if (hit) return hit;
    }
    return undefined;
  };
  const hit = walk(json);
  if (!hit) throw new Error("no link-marked run in parsed document");
  return hit;
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

  it("keeps a bare link mark un-styled (a TOC entry stays plain)", () => {
    // A link mark alone (no Hyperlink character style attr) compiles to a
    // hyperlink container whose runs carry no w:rStyle — Word's TOC entries
    // link without the link look. The style must come from the run's textStyle
    // attr (the insert paths stamp it), never from the container.
    const json = parseDOCX(generateDOCXSync(docWithLink("#toc-1")) as Uint8Array);
    const run = firstRunWithLink(json);
    expect(run.marks!.some((m) => m.type === "link")).toBe(true);
    const textStyle = run.marks!.find((m) => m.type === "textStyle");
    expect(textStyle?.attrs?.style ?? null).toBe(null);
  });

  it("round-trips the Hyperlink character style stamped on inserted links", () => {
    const json = parseDOCX(
      generateDOCXSync({
        ...docWithLink("https://example.com"),
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "text",
                text: "portal",
                marks: [
                  { type: "link", attrs: { href: "https://example.com", target: "_blank" } },
                  { type: "textStyle", attrs: { style: "Hyperlink" } },
                ],
              },
            ],
          },
        ],
      }) as Uint8Array,
    );
    const run = firstRunWithLink(json);
    const textStyle = run.marks!.find((m) => m.type === "textStyle");
    expect(textStyle?.attrs?.style).toBe("Hyperlink");
  });
});
