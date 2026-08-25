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
