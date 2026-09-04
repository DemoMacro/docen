import { Document, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { addSpellWord, checkSpelling, ignoreSpellWord, spellSuggestions } from "./spelling";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

/** A schema-only headless editor holding one text run per paragraph. */
const build = (...paragraphs: string[]): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text],
    content: {
      type: "doc",
      content: paragraphs.map((text) => ({
        type: "paragraph",
        content: [{ type: "text", text }],
      })),
    },
  });

describe("checkSpelling", () => {
  it("flags unknown words with their PM positions", () => {
    expect(checkSpelling(build("Hello zqqq world").state.doc)).toEqual([
      { from: 7, to: 11, word: "zqqq" },
    ]);
  });

  it("passes dictionary words, CJK runs, and numbers", () => {
    expect(checkSpelling(build("the quick 甲乙丙丁 123").state.doc)).toEqual([]);
  });

  it("lists every occurrence in document order", () => {
    // Paragraph 2's text starts at 1 + (9 + 2 open/close tokens) = 12.
    expect(checkSpelling(build("zqqq here", "more zqqq").state.doc).map((i) => i.from)).toEqual([
      1, 17,
    ]);
  });

  it("suppresses words added to the session dictionary", () => {
    addSpellWord("zqwww");
    expect(checkSpelling(build("zqwww now").state.doc)).toEqual([]);
  });

  it("skips a word ignored for the session", () => {
    ignoreSpellWord("qqzzw");
    expect(checkSpelling(build("a qqzzw b").state.doc)).toEqual([]);
  });
});

describe("spellSuggestions", () => {
  it("suggests close dictionary words, best first", () => {
    expect(spellSuggestions("speling")[0]).toBe("spelling");
  });

  it("returns nothing when the word is far from anything", () => {
    expect(spellSuggestions("zqqqqzzz")).toEqual([]);
  });
});
