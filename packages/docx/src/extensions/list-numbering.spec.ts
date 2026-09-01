import { generateDocumentSync } from "@office-open/docx";
import { getSchema } from "@tiptap/core";
import { parseHTML as parseLinkedomHTML } from "linkedom";
import { describe, expect, it } from "vitest";

import { compileDocument, parseDOCX } from "../converters/docx";
import { generateMarkdown, parseMarkdown } from "../converters/markdown";
import { docxExtensions, type JSONContent } from "../core";
import {
  assignOrderedReferences,
  buildListLevels,
  BULLET_REFERENCE,
  isGeneratedListReference,
  nextOrderedReference,
} from "./list-numbering";
import { parseHTMLBody } from "./paste";

describe("list-numbering builders", () => {
  it("builds a decimal ordered definition with per-level restarts", () => {
    const levels = buildListLevels("docen-ordered-1")!;
    expect(levels).toHaveLength(9);
    expect(levels[0]).toMatchObject({ level: 0, format: "decimal", start: 1, text: "%1." });
    // Deeper levels cascade their ancestors (Word's default multilevel shape),
    // so stepping a paragraph's level visibly renumbers it.
    expect(levels[8]).toMatchObject({ level: 8, text: "%1.%2.%3.%4.%5.%6.%7.%8.%9." });
  });

  it("maps the lower-roman variant to a level-0 lowerRoman format", () => {
    const levels = buildListLevels("docen-ordered-lower-roman-2")!;
    expect(levels[0]).toMatchObject({ format: "lowerRoman" });
    expect(levels[1]).toMatchObject({ format: "decimal" });
  });

  it("maps the circle variant to a circle glyph at every level", () => {
    const levels = buildListLevels("docen-bullet-circle")!;
    expect(levels[0]).toMatchObject({ format: "bullet", text: "○" });
    expect(levels[0].paragraph?.indent).toMatchObject({ left: 720, hanging: 360 });
  });

  it("recognizes only generated references", () => {
    expect(isGeneratedListReference(BULLET_REFERENCE)).toBe(true);
    expect(isGeneratedListReference("docen-ordered-3")).toBe(true);
    expect(isGeneratedListReference("list_3")).toBe(false);
  });
});

describe("nextOrderedReference", () => {
  it("numbers past the highest existing suffix", () => {
    expect(nextOrderedReference(["docen-ordered-1", "docen-ordered-3"], undefined)).toBe(
      "docen-ordered-4",
    );
  });

  it("shares the counter across variants and keeps the variant suffix", () => {
    expect(nextOrderedReference(["docen-ordered-lower-roman-2"], undefined, "lower-alpha")).toBe(
      "docen-ordered-lower-alpha-3",
    );
  });

  it("considers the numbering definitions too", () => {
    const numbering = { abstractNumberings: [{ reference: "docen-ordered-7" }] };
    expect(nextOrderedReference([], numbering)).toBe("docen-ordered-8");
  });
});

describe("assignOrderedReferences", () => {
  const para = (reference: string | null, level = 0): JSONContent =>
    reference
      ? ({
          type: "paragraph",
          attrs: { numbering: { reference, level } },
        } as JSONContent)
      : ({ type: "paragraph" } as JSONContent);

  it("splits placeholder runs into independent references", () => {
    const doc = {
      content: [para("html-ordered"), para("html-ordered"), para(null), para("html-ordered")],
    };
    assignOrderedReferences(doc);
    const refs = (doc.content as JSONContent[]).map(
      (n) => (n.attrs as { numbering?: { reference?: string } })?.numbering?.reference,
    );
    expect(refs).toEqual(["docen-ordered-1", "docen-ordered-1", undefined, "docen-ordered-2"]);
  });

  it("keeps deeper placeholder levels inside the same run", () => {
    const doc = {
      content: [para("html-ordered", 0), para("html-ordered", 1)],
    };
    assignOrderedReferences(doc);
    const refs = (doc.content as JSONContent[]).map(
      (n) => (n.attrs as { numbering: { reference: string } }).numbering.reference,
    );
    expect(refs[0]).toBe(refs[1]);
  });
});

describe("compile registers generated list definitions", () => {
  it("adds a decimal definition per referenced docen-ordered reference", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        { type: "paragraph", attrs: { numbering: { reference: "docen-ordered-1", level: 0 } } },
        { type: "paragraph", attrs: { numbering: { reference: "docen-ordered-2", level: 0 } } },
      ],
    };
    const opts = compileDocument(doc);
    const refs = (opts.numbering?.abstractNumberings ?? []).map((a) => a.reference);
    expect(refs).toContain("docen-ordered-1");
    expect(refs).toContain("docen-ordered-2");
  });

  it("round-trips a docx with generated list references", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
          content: [{ type: "text", text: "one" }],
        },
        {
          type: "paragraph",
          attrs: { bullet: { level: 1 } },
          content: [{ type: "text", text: "two" }],
        },
      ],
    };
    // The generated document carries the numbering part with both references;
    // the bullet sugar needs no definition (office-open's built-in numId 1).
    const bytes = generateDocumentSync(compileDocument(doc));
    expect(bytes.byteLength).toBeGreaterThan(0);
  });

  it("parses a real list docx back to flat list paragraphs (near-identity)", async () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
          content: [{ type: "text", text: "one" }],
        },
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 1 } },
          content: [{ type: "text", text: "two" }],
        },
      ],
    };
    const bytes = generateDocumentSync(compileDocument(doc));
    const parsed = parseDOCX(bytes);
    const content = parsed.content ?? [];
    expect(content).toHaveLength(2);
    expect(content[0].type).toBe("paragraph");
    // numId 1 is reserved for office-open's built-in default bullet list, so
    // the first caller-supplied definition lands on numId 2 → "list_2".
    expect(content[0].attrs?.numbering).toMatchObject({ reference: "list_2", level: 0 });
    expect(content[1].attrs?.numbering).toMatchObject({ reference: "list_2", level: 1 });
    // Both paragraphs share the source numId → one reference → the counter
    // continues across them (Word semantics for a shared concrete num).
  });
});

describe("markdown flat lists", () => {
  it("parses nested lists into leveled list paragraphs", () => {
    const json = parseMarkdown("- a\n    - b\n\n1. x\n2. y");
    const content = json.content ?? [];
    expect(content).toHaveLength(4);
    expect(content[0].attrs).toMatchObject({ bullet: { level: 0 } });
    expect(content[1].attrs).toMatchObject({ bullet: { level: 1 } });
    expect(content[2].attrs?.numbering).toMatchObject({ reference: "docen-ordered-1", level: 0 });
    expect(content[3].attrs?.numbering).toMatchObject({ reference: "docen-ordered-1", level: 0 });
  });

  it("serializes list paragraphs back as markdown lists", () => {
    const md = generateMarkdown({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { bullet: { level: 0 } },
          content: [{ type: "text", text: "a" }],
        },
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
          content: [{ type: "text", text: "x" }],
        },
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
          content: [{ type: "text", text: "y" }],
        },
      ],
    });
    expect(md).toContain("- a");
    expect(md).toContain("1. x");
    expect(md).toContain("2. y");
  });
});

describe("html flat lists", () => {
  it("parses nested ul/ol into leveled list paragraphs", () => {
    // Wrap in a full document: linkedom does not synthesize <body> for a bare
    // fragment, and document.body would be empty.
    const { document } = parseLinkedomHTML(
      `<!DOCTYPE html><html><body><ul><li>a<ul><li>b</li></ul></li><li>c</li></ul><ol><li>x</li></ol></body></html>`,
    );
    const json = parseHTMLBody(document.body as HTMLElement, getSchema(docxExtensions));
    const content = json.content ?? [];
    expect(content).toHaveLength(4);
    const levels = content.map((n) => n.attrs?.bullet?.level ?? n.attrs?.numbering?.level);
    expect(levels).toEqual([0, 1, 0, 0]);
  });
});
