import { describe, expect, it } from "vitest";

import { parseMarkdownTokens } from "./parse";

describe("parse IR shapes", () => {
  it("maps heading depth", () => {
    expect(parseMarkdownTokens("### Deep")).toEqual([
      { type: "heading", depth: 3, children: [{ type: "text", text: "Deep" }] },
    ]);
  });

  it("keeps list nesting in the IR and records ordered starts", () => {
    const blocks = parseMarkdownTokens("3. top\n    - inner");
    expect(blocks).toEqual([
      {
        type: "list",
        ordered: true,
        start: 3,
        items: [
          {
            blocks: [
              { type: "paragraph", children: [{ type: "text", text: "top" }] },
              {
                type: "list",
                ordered: false,
                items: [
                  { blocks: [{ type: "paragraph", children: [{ type: "text", text: "inner" }] }] },
                ],
              },
            ],
          },
        ],
      },
    ]);
  });

  it("marks the header row and column alignment of a table", () => {
    const [table] = parseMarkdownTokens("| A | B |\n| --- | :-: |\n| a | b |") as Extract<
      ReturnType<typeof parseMarkdownTokens>[number],
      { type: "table" }
    >[];
    if (table.type !== "table") throw new Error("not a table");
    expect(table.align).toEqual([null, "center"]);
    expect(table.rows[0].cells.every((c) => c.header)).toBe(true);
    expect(table.rows[1].cells.every((c) => !c.header)).toBe(true);
  });

  it("nests quotes and translates inline marks", () => {
    const [quote] = parseMarkdownTokens("> **bold** and `code`") as {
      type: "quote";
      blocks: unknown[];
    }[];
    if (quote.type !== "quote") throw new Error("not a quote");
    expect(quote.blocks).toHaveLength(1);
    expect(quote.blocks[0]).toMatchObject({
      type: "paragraph",
      children: [
        { type: "text", text: "bold", marks: [{ type: "bold" }] },
        { type: "text", text: " and " },
        { type: "text", text: "code", marks: [{ type: "code" }] },
      ],
    });
  });

  it("keeps block html and inline html as literal text", () => {
    expect(parseMarkdownTokens("<div>x</div>")).toEqual([
      { type: "paragraph", children: [{ type: "text", text: "<div>x</div>" }] },
    ]);
    const [p] = parseMarkdownTokens("a <b>bold</b> tag") as {
      type: "paragraph";
      children: { text: string }[];
    }[];
    if (p.type !== "paragraph") throw new Error("not a paragraph");
    expect(p.children.map((c) => ("text" in c ? c.text : "")).join("")).toBe("a <b>bold</b> tag");
  });

  it("records GFM task checkboxes on list items", () => {
    const [list] = parseMarkdownTokens("- [ ] todo\n- [x] done\n- plain") as Extract<
      ReturnType<typeof parseMarkdownTokens>[number],
      { type: "list" }
    >[];
    if (list.type !== "list") throw new Error("not a list");
    expect(list.items.map((i) => i.checked)).toEqual([false, true, undefined]);
  });

  it("lifts footnote definition lines into their own blocks", () => {
    const blocks = parseMarkdownTokens("[^1]: first **note**\n\nbody");
    expect(blocks).toEqual([
      {
        type: "footnoteDefinition",
        label: "1",
        children: [
          { type: "text", text: "first " },
          { type: "text", text: "note", marks: [{ type: "bold" }] },
        ],
      },
      { type: "paragraph", children: [{ type: "text", text: "body" }] },
    ]);
  });

  it("lifts single-word footnote bodies lexed as link reference definitions", () => {
    // marked lexes `[^1]: note` (no space in the destination) as a `def`
    // token, not a paragraph.
    expect(parseMarkdownTokens("[^1]: note")).toEqual([
      { type: "footnoteDefinition", label: "1", children: [{ type: "text", text: "note" }] },
    ]);
  });
});
