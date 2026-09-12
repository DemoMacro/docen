// @vitest-environment node
import { describe, expect, it } from "vitest";

import { Editor, Node as TextNode } from "../core";
import type { JSONContent } from "../index";
import { Document } from "./document";
import { PageBreak } from "./page-break";
import { Paragraph } from "./paragraph";

// The trailing-content rule is schema-agnostic (it only matches the
// "pageBreak" node name), so a minimal schema exercises it.
const Text = TextNode.create({ name: "text", group: "inline" });
const extensions = [Document, Paragraph, Text, PageBreak];

const editorWith = (paragraphs: JSONContent[][]): Editor => {
  const editor = new Editor({
    element: null,
    extensions,
    content: {
      type: "doc",
      content: paragraphs.map((children) => ({ type: "paragraph", content: children })),
    },
  });
  // element:null skips mount() and with it plugin installation — the
  // production bridge registers the extension plugins by hand (edit-bridge.ts).
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

/** Each top-level paragraph as a space-joined child-type list, e.g.
 *  "text pageBreak" — the split's structural fingerprint. */
const shapes = (editor: Editor): string[] =>
  editor.state.doc.content.content.map((p) => p.content.content.map((c) => c.type.name).join(" "));

describe("PageBreak trailing-content plugin", () => {
  it("splits the paragraph when content lands after a page break", () => {
    // A break at the paragraph end is legal; the paste then lands content
    // after it — the plugin must split (Word's break = paragraph terminator).
    const editor = editorWith([[{ type: "text", text: "before" }, { type: "pageBreak" }]]);
    editor.commands.command(({ tr }) => {
      tr.insertText("after", 8);
      return true;
    });
    expect(shapes(editor)).toEqual(["text pageBreak", "text"]);
  });

  it("re-splits when an undo joins the split back", () => {
    const editor = editorWith([[{ type: "text", text: "before" }, { type: "pageBreak" }]]);
    editor.commands.command(({ tr }) => {
      tr.insertText("after", 8);
      return true;
    });
    expect(shapes(editor)).toEqual(["text pageBreak", "text"]);
    // Simulate the undo join: merge the trailing paragraph (boundary at 9)
    // back into the break's paragraph. The join's changed range covers the
    // merged textblock, so the invariant is restored in the same dispatch.
    editor.commands.command(({ tr }) => {
      tr.join(9, 1);
      return true;
    });
    expect(shapes(editor)).toEqual(["text pageBreak", "text"]);
  });

  it("leaves a pre-existing mid-paragraph break alone on unrelated edits", () => {
    // A file loaded with `before<br page>after` in one paragraph renders
    // through the projection's block split; the plugin normalizes a paragraph
    // only when an edit touches it.
    const editor = editorWith([
      [{ type: "text", text: "before" }, { type: "pageBreak" }, { type: "text", text: "after" }],
      [{ type: "text", text: "other" }],
    ]);
    editor.commands.command(({ tr }) => {
      tr.insertText("X", 16);
      return true;
    });
    expect(editor.state.doc.content.content[0].childCount).toBe(3);
  });
});
