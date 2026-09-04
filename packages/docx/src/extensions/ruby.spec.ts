import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCX, type JSONContent } from "../index";

// Phonetic guide (拼音指南) round-trip: a w:ruby container resolves to text
// nodes carrying a ruby mark (annotation text + CT_RubyPr fields), and the
// mark compiles back into the container. The base run's other marks ride
// along; the annotation text stays flat on the mark.

const RUBY_ATTRS = {
  text: "jiǎ",
  alignment: "center",
  fontSize: 5,
  baseFontSize: 10,
  raise: 5,
  languageId: "zh-CN",
  dirty: null,
};

function docWithRuby(extraMarks: JSONContent["marks"] = []): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "甲乙",
            marks: [...extraMarks, { type: "ruby", attrs: { ...RUBY_ATTRS } }],
          },
        ],
      },
    ],
  };
}

/** The first text node's ruby mark (undefined = none). */
function rubyMark(json: JSONContent): Record<string, unknown> | undefined {
  for (const child of json.content ?? []) {
    if (child.type === "text") {
      const mark = (child.marks ?? []).find((m) => m.type === "ruby");
      if (mark) return mark.attrs as Record<string, unknown>;
    }
    const hit = rubyMark(child);
    if (hit) return hit;
  }
  return undefined;
}

describe("ruby (phonetic guide)", () => {
  it("round-trips a ruby mark through the DOCX container", () => {
    const json = parseDOCX(generateDOCXSync(docWithRuby()));
    const attrs = rubyMark(json);
    expect(attrs).toBeDefined();
    expect(attrs?.text).toBe("jiǎ");
    expect(attrs?.alignment).toBe("center");
    expect(attrs?.fontSize).toBe(5);
    expect(attrs?.baseFontSize).toBe(10);
    expect(attrs?.raise).toBe(5);
    expect(attrs?.languageId).toBe("zh-CN");
    // The base text is intact and editable — same text node shape as plain runs.
    const para = json.content?.[0];
    expect(para?.content?.[0]?.type).toBe("text");
    expect(para?.content?.[0]?.text).toBe("甲乙");
  });

  it("keeps sibling marks on the base run across the container", () => {
    const json = parseDOCX(generateDOCXSync(docWithRuby([{ type: "bold", attrs: {} }])));
    const types = (json.content?.[0]?.content?.[0]?.marks ?? []).map((m) => m.type);
    // (the compile pipeline also stamps a textStyle carrier — ignore it)
    expect(types).toContain("bold");
    expect(types).toContain("ruby");
  });

  it("carries Word's recalculate (dirty) flag through the container", () => {
    const doc = docWithRuby();
    const marks = doc.content?.[0]?.content?.[0]?.marks ?? [];
    const attrs = marks.find((m) => m.type === "ruby")?.attrs as Record<string, unknown>;
    attrs.dirty = true;
    const json = parseDOCX(generateDOCXSync(doc));
    expect(rubyMark(json)?.dirty).toBe(true);
  });

  it("leaves plain runs untouched", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "无注音" },
            { type: "text", text: "甲", marks: [{ type: "ruby", attrs: { ...RUBY_ATTRS } }] },
          ],
        },
      ],
    };
    const json = parseDOCX(generateDOCXSync(doc));
    const texts = json.content?.[0]?.content ?? [];
    expect(texts).toHaveLength(2);
    expect(texts[0]?.text).toBe("无注音");
    expect((texts[0]?.marks ?? []).some((m) => m.type === "ruby")).toBe(false);
    expect((texts[1]?.marks ?? []).some((m) => m.type === "ruby")).toBe(true);
  });
});
