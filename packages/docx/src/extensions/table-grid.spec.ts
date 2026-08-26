import { generateDocumentSync } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, parseDOCX } from "../converters/docx";
import { projectDocumentOptions } from "../layout/project";
import { parseDocxBlock, parseDocx, renderDocx } from "./table";
import { parseDocx as parseCellDocx, renderDocx as renderCellDocx } from "./table-cell";
import type { ResolveContext } from "./types";

const ctx = {
  parseNodeAttrs: (_kind: string, opts: unknown) => opts,
  resolveBlockStream: (children: unknown[]) =>
    (children as { paragraph: unknown }[]).map((c) => ({ type: "paragraph", content: [] })),
  styles: undefined,
} as unknown as ResolveContext;

describe("flat table grid", () => {
  it("passes cell structural attrs through verbatim (columnSpan/verticalMerge)", () => {
    const attrs = parseCellDocx({
      columnSpan: 2,
      verticalMerge: "restart",
      width: { size: 1440, type: "twips" },
      children: [],
    } as never);
    expect(attrs.columnSpan).toBe(2);
    expect(attrs.verticalMerge).toBe("restart");
    expect(attrs.width).toEqual({ size: 1440, type: "twips" });
    const back = renderCellDocx({ type: "tableCell", attrs } as never);
    expect(back.columnSpan).toBe(2);
    expect(back.verticalMerge).toBe("restart");
  });

  it("keeps vMerge cells as real nodes (no rowspan rebuild)", () => {
    const node = parseDocxBlock.convert(
      {
        table: {
          rows: [
            {
              cells: [{ verticalMerge: "restart", children: [] }, { children: [] }],
            },
            {
              cells: [{ verticalMerge: "continue", columnSpan: 1, children: [] }, { children: [] }],
            },
          ],
        },
      },
      ctx,
    );
    const rows = (node?.content ?? []) as { content: unknown[] }[];
    expect(rows).toHaveLength(2);
    // Row 0's restart and row 1's continue are both real tableCell nodes.
    expect((rows[1].content[0] as { attrs?: Record<string, unknown> }).attrs?.verticalMerge).toBe(
      "continue",
    );
  });

  it("round-trips a vMerge table through the real docx pipeline", () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "table",
          content: [
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  attrs: { verticalMerge: "restart" },
                  content: [{ type: "paragraph" }],
                },
                { type: "tableCell", content: [{ type: "paragraph" }] },
              ],
            },
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  attrs: { verticalMerge: "continue" },
                  content: [{ type: "paragraph" }],
                },
                { type: "tableCell", content: [{ type: "paragraph" }] },
              ],
            },
          ],
        },
      ],
    };
    const bytes = generateDocumentSync(compileDocument(doc as never));
    const parsed = parseDOCX(bytes);
    const table = parsed.content?.[0];
    expect(table?.type).toBe("table");
    const rows = (table?.content ?? []) as { content: { attrs?: Record<string, unknown> }[] }[];
    expect(rows[0].content[0].attrs?.verticalMerge).toBe("restart");
    expect(rows[1].content[0].attrs?.verticalMerge).toBe("continue");
  });

  it("counts grid columns from columnSpan (a spanning cell covers its columns)", () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "table",
          attrs: { columnWidths: [2000, 3000, 4000] },
          content: [
            {
              type: "tableRow",
              content: [
                { type: "tableCell", attrs: { columnSpan: 2 }, content: [{ type: "paragraph" }] },
                { type: "tableCell", content: [{ type: "paragraph" }] },
              ],
            },
          ],
        },
      ],
    };
    const bytes = generateDocumentSync(compileDocument(doc as never));
    const parsed = parseDOCX(bytes);
    const table = parsed.content?.[0];
    // A 2+1 first row spans 3 grid columns — the tblGrid must keep all three.
    expect(((table?.attrs?.columnWidths as number[]) ?? []).length).toBe(3);
  });

  it("expands vMerge into rowspan at the single projection point", () => {
    const { blocks } = projectDocumentOptions({
      sections: [
        {
          children: [
            {
              table: {
                rows: [
                  {
                    cells: [{ verticalMerge: "restart", children: [] }, { children: [] }],
                  },
                  {
                    cells: [{ verticalMerge: "continue", children: [] }, { children: [] }],
                  },
                  {
                    cells: [{ verticalMerge: "continue", children: [] }, { children: [] }],
                  },
                ] as never[],
              },
            },
          ],
        },
      ],
    } as never);
    const table = blocks[0];
    expect(table?.kind).toBe("table");
    const rows = (table as unknown as { rows: { cells: { rowspan: number }[] }[] }).rows;
    // Continuation rows drop the merged cell; the restart carries rowspan 3.
    expect(rows).toHaveLength(3);
    expect(rows[0].cells[0].rowspan).toBe(3);
    expect(rows[1].cells).toHaveLength(1);
    expect(rows[2].cells).toHaveLength(1);
  });
});
