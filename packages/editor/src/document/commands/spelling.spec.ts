import { Document, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import type { SpellingIssue } from "../spelling";
import { SpellingCommands, type SpellingHost } from "./spelling";

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

/** A host recording what the bridge was handed — the squiggle feed. */
const hostOf = (editor: EditorType): { host: SpellingHost; pushed: SpellingIssue[][] } => {
  const pushed: SpellingIssue[][] = [];
  return {
    host: {
      editor: () => editor,
      bridge: () => ({
        setSpellingIssues(issues: SpellingIssue[]): void {
          pushed.push(issues);
        },
        scrollIntoView(): void {},
      }),
      // No shadow DOM in specs — run()'s status-bar write is a no-op there.
      element: () => ({ shadowRoot: null }) as HTMLElement,
    },
    pushed,
  };
};

describe("SpellingCommands.mapThrough", () => {
  it("shifts the issue ranges through an edit before the word", () => {
    const editor = build("Hello zqqq world");
    const { host, pushed } = hostOf(editor);
    const cmd = new SpellingCommands(host);
    cmd.run();
    expect(cmd.issues()).toEqual([{ from: 7, to: 11, word: "zqqq" }]);

    // Two characters typed at the paragraph's head — the word slides right
    // by 2. (pos 0 is the doc/block boundary: inserting there opens a new
    // paragraph, a +4 shift — not the typing case under test.)
    cmd.mapThrough(editor.state.tr.insertText("xx", 1));
    expect(cmd.issues()).toEqual([{ from: 9, to: 13, word: "zqqq" }]);
    // The carried ranges are pushed so the next render can draw them.
    expect(pushed.at(-1)).toEqual([{ from: 9, to: 13, word: "zqqq" }]);
  });

  it("drops a word the edit deleted outright", () => {
    const editor = build("Hello zqqq world");
    const { host } = hostOf(editor);
    const cmd = new SpellingCommands(host);
    cmd.run();

    cmd.mapThrough(editor.state.tr.delete(7, 11));
    expect(cmd.issues()).toEqual([]);
  });
});

describe("SpellingCommands.setEnabled", () => {
  it("off clears the issues and the squiggle feed; on re-checks", () => {
    const editor = build("Hello zqqq world");
    const { host, pushed } = hostOf(editor);
    const cmd = new SpellingCommands(host);
    cmd.run();
    expect(cmd.issues()).toHaveLength(1);

    cmd.setEnabled(false);
    expect(cmd.issues()).toEqual([]);
    expect(pushed.at(-1)).toEqual([]);

    cmd.setEnabled(true);
    expect(cmd.issues()).toEqual([{ from: 7, to: 11, word: "zqqq" }]);
  });
});
