import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateMarkdown, parseMarkdown } from "./markdown";

describe("markdown heading round-trip", () => {
  it("serializes a heading paragraph with # prefix and parses ## back", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { heading: "Heading2" },
          content: [{ type: "text", text: "chapter" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "hello" }] },
      ],
    };
    const md = generateMarkdown(doc);
    expect(md).toContain("## chapter");
    expect(md).toContain("hello");

    const parsed = parseMarkdown("## A title\n\nbody text");
    const head = parsed.content?.[0] as { type: string; attrs?: { heading?: string } };
    expect(head.type).toBe("paragraph");
    expect(head.attrs?.heading).toBe("Heading2");
  });
});

describe("markdown information preservation", () => {
  it("keeps the code fence language in a docen-only attr", () => {
    const parsed = parseMarkdown("```ts\nconst a = 1;\n```");
    const code = parsed.content?.[0] as { attrs?: { style?: string; codeLanguage?: string } };
    expect(code.attrs?.style).toBe("Code");
    expect(code.attrs?.codeLanguage).toBe("ts");
    expect(generateMarkdown(parsed)).toBe("```ts\nconst a = 1;\n```");
  });

  it("renders task items as plain bullets with a checkbox text prefix", () => {
    const parsed = parseMarkdown("- [ ] todo\n- [x] done");
    const first = parsed.content?.[0] as {
      attrs?: { bullet?: { level?: number } };
      content?: { text?: string }[];
    };
    // No task semantics in DOCX — the checkbox becomes literal text.
    expect(first.attrs?.bullet?.level).toBe(0);
    expect(first.content?.[0]?.text).toBe("[ ] ");
    // Exporting lifts the prefix back to GFM task syntax, unescaped.
    expect(generateMarkdown(parsed)).toBe("- [ ] todo\n- [x] done");
  });

  it("keeps a custom ordered start in the paragraph and the abstract definition", () => {
    const parsed = parseMarkdown("3. third\n4. fourth");
    const first = parsed.content?.[0] as {
      attrs?: { numbering?: { start?: number; level?: number } };
    };
    expect(first.attrs?.numbering?.start).toBe(3);
    const doc = parsed.attrs as {
      numbering?: {
        abstractNumberings?: { levels?: { level: number; start: number }[] }[];
      };
    };
    expect(doc.numbering?.abstractNumberings?.[0]?.levels?.[0]).toMatchObject({
      level: 0,
      start: 3,
    });
    expect(generateMarkdown(parsed)).toBe("3. third\n4. fourth");
  });

  it("appends documentExtras footnotes at the end as definitions", () => {
    const doc: JSONContent = {
      type: "doc",
      attrs: {
        documentExtras: {
          footnotes: [
            { id: 1, children: [{ type: "paragraph", content: [{ type: "text", text: "note" }] }] },
          ],
        },
      },
      content: [{ type: "paragraph", content: [{ type: "text", text: "body" }] }],
    };
    const md = generateMarkdown(doc);
    expect(md).toBe("body\n\n[^1]: note");
    // The definition re-imports as a plain footnote paragraph and stabilizes.
    expect(generateMarkdown(parseMarkdown(md))).toBe(md);
  });
});

describe("markdown mapper round-trip", () => {
  it("reaches a stable fixpoint for a mixed document", () => {
    const md = [
      "# Title",
      "",
      "A **bold** and *italic* line with `code`.",
      "",
      "- alpha",
      "    - nested",
      "",
      "1. first",
      "2. second",
      "",
      "> quoted line",
      "",
      "| h1 | h2 |",
      "| --- | :-: |",
      "| a | b |",
      "",
      "---",
    ].join("\n");
    const parsed = parseMarkdown(md);
    const head = parsed.content?.[0] as { type: string; attrs?: { heading?: string } };
    expect(head.attrs?.heading).toBe("Heading1");
    const bullet = parsed.content?.find((n) => n.attrs?.bullet) as {
      attrs?: { bullet?: { level?: number } };
    };
    expect(bullet.attrs?.bullet?.level).toBe(0);

    const once = generateMarkdown(parsed);
    const twice = generateMarkdown(parseMarkdown(once));
    expect(twice).toBe(once);
    expect(once).toContain("    - nested");
    expect(once).toContain("1. first");
    expect(once).toContain("> quoted line");
    expect(once).toContain("| --- | :-: |");
  });
});
