import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCX, type JSONContent } from "../index";

// The Review tab's comment commands anchor a comment by stamping
// commentRangeStart/commentRangeEnd/commentReference markers into the body
// (inlinePassthrough atoms in the editor model) while the content rides
// doc.attrs.documentExtras.comments. This pins the whole channel — the exact
// JSON an inserted comment produces must generate a DOCX that parses back
// with the structured comment (author/date/content) and the range markers
// intact, ready for the editor to re-anchor them.

function marker(data: object): JSONContent {
  return { type: "inlinePassthrough", attrs: { data: JSON.stringify(data) } };
}

const INSERTED_DOC: JSONContent = {
  type: "doc",
  attrs: {
    documentExtras: {
      comments: [
        {
          id: 1,
          author: "Reviewer",
          initials: "R",
          date: "2026-08-28T00:00:00Z",
          children: [{ text: "Please expand this." }],
        },
      ],
    },
  },
  content: [
    {
      type: "paragraph",
      content: [
        marker({ commentRangeStart: { id: 1 } }),
        { type: "text", text: "anchored text" },
        marker({ commentRangeEnd: { id: 1 } }),
        marker({ commentReference: 1 }),
      ],
    },
  ],
};

function walk(node: JSONContent, visit: (n: JSONContent) => void): void {
  visit(node);
  for (const child of node.content ?? []) walk(child, visit);
}

describe("comment round-trip", () => {
  it("keeps the structured comment and the body range markers", () => {
    const json = parseDOCX(generateDOCXSync(INSERTED_DOC) as Uint8Array);
    const attrs = (json.attrs ?? {}) as { documentExtras?: { comments?: Record<string, any>[] } };
    const comments = attrs.documentExtras?.comments ?? [];
    expect(comments).toHaveLength(1);
    expect(comments[0].author).toBe("Reviewer");
    expect(comments[0].date).toBe("2026-08-28T00:00:00Z");
    expect(JSON.stringify(comments[0].children)).toContain("Please expand this.");

    const data: string[] = [];
    walk(json, (n) => {
      if (n.type === "inlinePassthrough") data.push(String(n.attrs?.data ?? ""));
    });
    expect(data.some((d) => d.includes("commentRangeStart"))).toBe(true);
    expect(data.some((d) => d.includes("commentRangeEnd"))).toBe(true);
    expect(data.some((d) => d.includes("commentReference"))).toBe(true);
  });
});
