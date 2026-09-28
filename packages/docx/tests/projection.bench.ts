// Layout-projection benchmark: a medium document (headings, formatted runs,
// a table) projected through the editor's per-transaction chain —
// Tiptap JSON → compileDocument → projectDocumentOptions — and the projection
// alone. Run with `pnpm exec vp test bench`. The fixture is built in code so
// the bench stays stable and needs no I/O.

import type { JSONContent } from "@tiptap/core";
import { bench, describe } from "vitest";

import { compileDocument } from "../src";
import { projectDocument } from "../src/layout";

const styles = {
  paragraphStyles: [
    {
      id: "Normal",
      default: true,
      paragraph: { spacing: { line: 360, lineRule: "auto" } },
      run: { size: 21, font: { ascii: "Times New Roman", eastAsia: "SimSun" } },
    },
    {
      id: "Heading1",
      basedOn: "Normal",
      paragraph: { spacing: { before: 240, after: 120 }, keepNext: true },
      run: { size: 32, bold: true, color: "2F5496" },
    },
  ],
  default: { document: { paragraph: { spacing: { after: 160 } }, run: { size: 21 } } },
};

const sectionProperties = {
  pageSize: { width: 11906, height: 16838 },
  pageMargin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
};

/** 40 headings + formatted paragraphs + a 6×3 table — a medium clause body. */
function buildMediumDocument(): JSONContent {
  const content: JSONContent[] = [];
  for (let i = 1; i <= 40; i++) {
    content.push({
      type: "paragraph",
      attrs: { heading: "Heading1" },
      content: [{ type: "text", text: `第 ${i} 条` }],
    });
    content.push({
      type: "paragraph",
      attrs: i % 2 === 0 ? { alignment: "both", indent: { firstLine: 480 } } : {},
      content: [
        { type: "text", text: "本条", marks: [{ type: "bold" }] },
        { type: "text", text: "规定", marks: [{ type: "italic" }] },
        { type: "text", text: "了排版行为，", marks: [{ type: "underline" }] },
        {
          type: "text",
          text: `第 ${i} 款的正文内容与格式。`,
          marks: [{ type: "textStyle", attrs: { color: "C00000", size: 24 } }],
        },
      ],
    });
  }
  content.push({
    type: "table",
    attrs: {
      columnWidths: [2400, 2400, 2400],
      width: { size: 100, type: "percent" },
    },
    content: Array.from({ length: 6 }, (_, row) => ({
      type: "tableRow",
      attrs: row === 0 ? { tableHeader: true, cantSplit: true } : {},
      content: Array.from({ length: 3 }, (_, col) => ({
        type: "tableCell",
        attrs: col === 0 ? { columnSpan: 1 } : {},
        content: [
          {
            type: "paragraph",
            content: [{ type: "text", text: `R${row + 1}C${col + 1}` }],
          },
        ],
      })),
    })),
  });
  return { type: "doc", attrs: { styles, sectionProperties }, content };
}

const model = buildMediumDocument();
const docOpts = compileDocument(model);

describe("layout projection", () => {
  bench("compileDocument + projectDocument (editor transaction)", () => {
    projectDocument(compileDocument(model));
  });

  bench("projectDocument (projection only)", () => {
    projectDocument(docOpts);
  });
});
