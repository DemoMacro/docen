import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { getSchema } from "@tiptap/core";
import { parseHTML as parseLinkedomHTML } from "linkedom";
import { describe, expect, it } from "vitest";

import { docxExtensions } from "../core";
import { parseHTMLBody } from "./paste";

const schema = getSchema(docxExtensions);

const parse = (html: string) => {
  const { document } = parseLinkedomHTML(`<!DOCTYPE html><html><body>${html}</body></html>`);
  return parseHTMLBody(document.body as HTMLElement, schema);
};

const firstTextMarks = (json: ReturnType<typeof parse>) => {
  const para = json.content?.[0];
  return para?.content?.[0]?.marks ?? [];
};

describe("paste style mapping", () => {
  it("maps <strong>/<b> to the bold mark", () => {
    expect(firstTextMarks(parse("<p><strong>bold</strong></p>"))).toContainEqual({ type: "bold" });
    expect(firstTextMarks(parse("<p><b>bold</b></p>"))).toContainEqual({ type: "bold" });
  });

  it("maps <em>/<i> to the italic mark and <u> to underline", () => {
    expect(firstTextMarks(parse("<p><em>x</em></p>"))).toContainEqual({ type: "italic" });
    expect(firstTextMarks(parse("<p><i>x</i></p>"))).toContainEqual({ type: "italic" });
    expect(firstTextMarks(parse("<p><u>x</u></p>"))).toContainEqual({ type: "underline" });
  });

  it("maps <s>/<del> to the strike mark", () => {
    expect(firstTextMarks(parse("<p><s>x</s></p>"))).toContainEqual({ type: "strike" });
    expect(firstTextMarks(parse("<p><del>x</del></p>"))).toContainEqual({ type: "strike" });
  });

  it("maps <h1>-<h6> to paragraphs with a HeadingN style attr", () => {
    const json = parse("<h1>Title</h1>");
    expect(json.content?.[0]?.type).toBe("paragraph");
    expect(json.content?.[0]?.attrs?.heading).toMatch(/^Heading\d$/);
  });

  it("maps <ul> to bullet paragraphs and <ol> to generated numbering references", () => {
    const json = parse("<ul><li>a</li></ul><ol><li>x</li></ol>");
    expect(json.content?.[0]?.attrs?.bullet).toMatchObject({ level: 0 });
    expect(json.content?.[1]?.attrs?.numbering?.reference).toMatch(/^docen-ordered-\d+$/);
  });

  it("maps <sub>/<sup> and <mark> to their marks", () => {
    expect(firstTextMarks(parse("<p><sub>x</sub></p>"))).toContainEqual({ type: "subscript" });
    expect(firstTextMarks(parse("<p><sup>x</sup></p>"))).toContainEqual({ type: "superscript" });
    // Highlight color comes from data-color or inline background-color — a
    // bare <mark> carries neither (DOM parsing reads no stylesheets), so it
    // maps to a colorless highlight.
    expect(
      firstTextMarks(parse(`<p><mark style="background-color: yellow">x</mark></p>`)),
    ).toContainEqual({ type: "highlight", attrs: { color: "yellow" } });
    expect(firstTextMarks(parse("<p><mark>x</mark></p>"))).toContainEqual({
      type: "highlight",
      attrs: { color: null },
    });
  });
});

describe("fixture corpus smoke (tests/html)", () => {
  it("parses every fixture without throwing", () => {
    const dir = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "tests", "html");
    const files = readdirSync(dir).filter((name) => name.endsWith(".html"));
    expect(files.length).toBeGreaterThanOrEqual(30);
    for (const name of files) {
      const json = parse(readFileSync(join(dir, name), "utf-8"));
      expect(json.type).toBe("doc");
    }
  });
});
