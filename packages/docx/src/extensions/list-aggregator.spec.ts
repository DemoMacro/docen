import type { DocumentOptions } from "@office-open/docx";
import { describe, it, expect } from "vitest";

import { DocxManager } from "../converters/docx";
import { docxExtensions } from "../core";

// buildListTree rebuilds a nested Tiptap list tree from a flat run of DOCX
// numbering paragraphs. The pipeline smoke suites (tests/docx.ts, tests/html.ts)
// only check that the round-trip doesn't throw, so a wrong tree would pass —
// these assertions pin the stack algorithm's nesting, start, and split behavior.

const mgr = new DocxManager(docxExtensions);

interface Item {
  text: string;
  depth: number;
}

/** Flatten a resolved doc into (text, list-nesting-depth) pairs. depth = the
 *  list nesting level an item sits at (top-level list → 0). */
function walk(nodes: unknown[] | undefined, depth: number, out: Item[]): void {
  for (const n of nodes ?? []) {
    if (typeof n !== "object" || !n) continue;
    const node = n as { type?: string; content?: unknown[] };
    if (node.type && node.type.endsWith("List")) {
      walk(node.content as unknown[], depth + 1, out);
    } else if (node.type && node.type.endsWith("Item")) {
      const para = (node.content ?? []).find(
        (c) => typeof c === "object" && c !== null && (c as { type?: string }).type === "paragraph",
      ) as { content?: { text?: string }[] } | undefined;
      const text = (para?.content ?? [])
        .map((t) =>
          typeof t === "object" && t !== null ? ((t as { text?: string }).text ?? "") : "",
        )
        .join("");
      out.push({ text, depth });
      walk(
        (node.content ?? []).filter(
          (c) =>
            typeof c === "object" &&
            c !== null &&
            ((c as { type?: string }).type ?? "").endsWith("List"),
        ),
        depth,
        out,
      );
    }
  }
}

function itemsOf(docOpts: DocumentOptions): Item[] {
  const json = mgr.resolve(docOpts);
  const out: Item[] = [];
  walk(json.content as unknown[], -1, out);
  return out;
}

/** A bullet list whose items sit at the given numbering levels. detectList reads
 *  levels[0].format to classify kind; numbering.level carries the depth. */
function bulletDoc(levels: number[]): DocumentOptions {
  return {
    sections: [
      {
        children: levels.map((level, i) => ({
          paragraph: {
            numbering: { reference: "L", level },
            children: [{ text: `i${i}` }],
          },
        })),
      },
    ],
    numbering: {
      config: [
        {
          reference: "L",
          levels: [{ format: "bullet" }, { format: "bullet" }, { format: "bullet" }],
        },
      ],
    },
  } as unknown as DocumentOptions;
}

describe("list-aggregator buildListTree", () => {
  it("nests by numbering level then returns to top level (0,1,2,0)", () => {
    const items = itemsOf(bulletDoc([0, 1, 2, 0]));
    expect(items.map((x) => x.depth)).toEqual([0, 1, 2, 0]);
    expect(items.map((x) => x.text)).toEqual(["i0", "i1", "i2", "i3"]);
  });

  it("keeps flat siblings in one list at level 0", () => {
    expect(itemsOf(bulletDoc([0, 0, 0])).map((x) => x.depth)).toEqual([0, 0, 0]);
  });

  it("carries start on an ordered list (format=decimal, start=3)", () => {
    const json = mgr.resolve({
      sections: [
        {
          children: [
            { paragraph: { numbering: { reference: "O", level: 0 }, children: [{ text: "a" }] } },
            { paragraph: { numbering: { reference: "O", level: 0 }, children: [{ text: "b" }] } },
          ],
        },
      ],
      numbering: { config: [{ reference: "O", levels: [{ format: "decimal", start: 3 }] }] },
    } as unknown as DocumentOptions);
    const first = (json.content as unknown as { type?: string; attrs?: { start?: number } }[])[0];
    expect(first?.type).toBe("orderedList");
    expect(first?.attrs?.start).toBe(3);
  });

  it("splits adjacent same-level lists with distinct references", () => {
    const json = mgr.resolve({
      sections: [
        {
          children: [
            { paragraph: { numbering: { reference: "A", level: 0 }, children: [{ text: "a" }] } },
            { paragraph: { numbering: { reference: "B", level: 0 }, children: [{ text: "b" }] } },
          ],
        },
      ],
      numbering: {
        config: [
          { reference: "A", levels: [{ format: "bullet" }] },
          { reference: "B", levels: [{ format: "bullet" }] },
        ],
      },
    } as unknown as DocumentOptions);
    const topLists = ((json.content as unknown[]) ?? []).filter(
      (n) =>
        typeof n === "object" &&
        n !== null &&
        ((n as { type?: string }).type ?? "").endsWith("List"),
    );
    expect(topLists).toHaveLength(2);
  });
});
