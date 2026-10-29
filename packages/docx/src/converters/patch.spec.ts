import { generateDocumentSync } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, patchDOCX, type JSONContent } from "../index";

const template = (text: string): Uint8Array =>
  generateDOCXSync({
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text }] }],
  }) as Uint8Array;

const replacement = (...paragraphs: readonly string[]): JSONContent => ({
  type: "doc",
  content: paragraphs.map((text) => ({
    type: "paragraph",
    content: [{ type: "text", text }],
  })),
});

const tableTemplate = (cells: readonly string[][]): Uint8Array =>
  generateDOCXSync({
    type: "doc",
    content: [
      {
        type: "table",
        content: cells.map((row) => ({
          type: "tableRow",
          content: row.map((text) => ({
            type: "tableCell",
            content: [{ type: "paragraph", content: [{ type: "text", text }] }],
          })),
        })),
      },
    ],
  });

describe("patchDOCX", () => {
  it("replaces placeholders in default headers", async () => {
    const source: DocumentOptions = {
      sections: [
        {
          headers: { default: [{ paragraph: { text: "Owner {{name}}" } }] },
          children: [],
        },
      ],
    };
    const patched = await patchDOCX({
      template: generateDocumentSync(source) as Uint8Array,
      patches: { name: { content: replacement("Ada") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);

    expect(JSON.stringify(document)).toContain("Owner Ada");
    expect(JSON.stringify(document)).not.toContain("{{name}}");
  });

  it("replaces an inline placeholder without splitting the sentence", async () => {
    const patched = await patchDOCX({
      template: template("Hello {{name}}"),
      patches: { name: { content: replacement("Ada") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);

    expect(JSON.stringify(document)).toContain("Hello Ada");
    expect(JSON.stringify(document)).not.toContain("{{name}}");
  });

  it("replaces a whole placeholder paragraph with multiple block children", async () => {
    const patched = await patchDOCX({
      template: template("{{body}}"),
      patches: { body: { content: replacement("First", "Second") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);

    expect(JSON.stringify(document)).toContain("First");
    expect(JSON.stringify(document)).toContain("Second");
    expect(JSON.stringify(document)).not.toContain("{{body}}");
  });

  it("replaces a placeholder inside a table cell without removing the table", async () => {
    const patched = await patchDOCX({
      template: tableTemplate([
        ["Keep", "{{body}}"],
        ["Left", "Right"],
      ]),
      patches: { body: { content: replacement("Replaced") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);
    const table = document.content?.find((node) => node.type === "table");
    const cell = (row: number, column: number) =>
      table?.content?.[row]?.content?.[column]?.content?.[0]?.content?.[0]?.text;

    expect(cell(0, 0)).toBe("Keep");
    expect(cell(0, 1)).toBe("Replaced");
    expect(cell(1, 0)).toBe("Left");
    expect(cell(1, 1)).toBe("Right");
  });

  it("inserts multiple paragraphs inside a table cell", async () => {
    const patched = await patchDOCX({
      template: tableTemplate([["Keep", "{{body}}"]]),
      patches: { body: { content: replacement("First", "Second") } },
      outputType: "uint8array",
    });
    const document = parseDOCXSync(patched);
    const table = document.content?.find((node) => node.type === "table");
    const cellParagraphs = table?.content?.[0]?.content?.[1]?.content;

    expect(cellParagraphs).toHaveLength(2);
    expect(JSON.stringify(cellParagraphs)).toContain("First");
    expect(JSON.stringify(cellParagraphs)).toContain("Second");
    expect(JSON.stringify(document)).not.toContain("{{body}}");
  });
});
