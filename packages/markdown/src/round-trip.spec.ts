import { describe, expect, it } from "vitest";

import { parseMarkdownTokens } from "./parse";
import { renderMarkdownBlocks } from "./render";

const roundTrip = (md: string): string => renderMarkdownBlocks(parseMarkdownTokens(md));

describe("markdown round-trip", () => {
  const cases: string[] = [
    "# Title",
    "## Heading two",
    "###### Deep heading",
    "A plain paragraph.",
    "Body with **bold** and *italic* and ~~strike~~ and `inline code`.",
    "**bold *nested italic* tail**",
    "*outer **inner** end*",
    "[Example](https://example.com)",
    '[titled](https://example.com "the title")',
    "![alt text](https://example.com/a.png)",
    "line one  \nline two",
    "- one\n- two\n- three",
    "- one\n    - nested\n    - nested two\n- back",
    "1. first\n2. second",
    "1. first\n    1. inner\n2. second",
    "3. third\n4. fourth",
    "> a quoted line\n> and another",
    "> para one\n>\n> para two",
    "> - listed\n> - items",
    "---",
    "```js\nconst a = 1;\n```",
    "```\nplain fence\n```",
    "| A | B |\n| --- | :-: |\n| a | b |",
    "| Left | Center | Right |\n| :--- | :-: | ---: |\n| 1 | 2 | 3 |",
    "# Heading\n\nParagraph.\n\n- item\n\n> quote\n\n---",
    "- [ ] open\n- [x] closed",
    "1. [x] first done\n2. [ ] second open",
    "[^1]: a footnote\n\nBody text.",
  ];

  it.each(cases)("round-trips: %j", (md) => {
    expect(roundTrip(md)).toBe(md);
  });

  it("keeps text intact through escape round-trip", () => {
    const md = "A \\*literal\\* line with \\[brackets\\] and a\\\\slash.";
    expect(roundTrip(md)).toBe(md);
  });

  it("splits runs of one emphasis group without broken delimiters", () => {
    const out = roundTrip("plain **bold *infix* bold again** end");
    expect(out).toBe("plain **bold *infix* bold again** end");
  });

  it("keeps one emphasis group across runs with minimal delimiters", () => {
    const blocks = parseMarkdownTokens("x");
    blocks.push({
      type: "paragraph",
      children: [
        { type: "text", text: "bold ", marks: [{ type: "bold" }] },
        { type: "text", text: "next", marks: [{ type: "bold" }, { type: "italic" }] },
      ],
    });
    const out = renderMarkdownBlocks(blocks);
    expect(out).toBe("x\n\n**bold *next***");
    expect(roundTrip(out)).toBe(out);
  });

  it("pads ragged table rows to a rectangular grid", () => {
    const out = renderMarkdownBlocks([
      {
        type: "table",
        align: [null, null],
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", children: [{ type: "text", text: "a" }] }] }] },
          {
            cells: [
              { blocks: [{ type: "paragraph", children: [{ type: "text", text: "b" }] }] },
              { blocks: [] },
            ],
          },
        ],
      },
    ]);
    expect(out).toBe("| a |   |\n| --- | --- |\n| b |   |");
    expect(roundTrip(out)).toBe(out);
  });
});
