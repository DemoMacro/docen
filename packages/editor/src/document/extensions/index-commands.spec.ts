import { Document, InlinePassthrough, Paragraph, Tab } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { IndexCommands } from "./index-commands";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

/** A body paragraph whose runs end with the given inline atoms. */
const body = (...inline: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "paragraph",
  content: inline,
});

const text = (value: string): Record<string, unknown> => ({ type: "text", text: value });

/** Word's invisible index marker: a cached-less XE fldSimple in the inline
 *  passthrough atom. */
const xe = (entry: string): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify({ simpleField: { instruction: `XE "${entry}"` } }) },
});

const docOf = (...blocks: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "doc",
  content: blocks,
});

/** A headless editor with exactly the schema the index workflow touches. */
const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Tab, InlinePassthrough, IndexCommands],
    content: doc,
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

interface IndexEntryInfo {
  style: unknown;
  indent: unknown;
  tabStops: unknown;
  text: string;
  pages: string;
}

/** Flatten the document's Index-styled paragraphs into testable shapes. */
const entriesOf = (editor: EditorType): IndexEntryInfo[] => {
  const out: IndexEntryInfo[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    const style = node.attrs.style;
    if (typeof style !== "string" || !/^Index\d$/.test(style)) return true;
    let entry = "";
    let pages = "";
    let inPages = false;
    node.content.forEach((child) => {
      if (child.type.name === "text") {
        if (inPages) pages += child.text ?? "";
        else entry += child.text ?? "";
      } else if (child.type.name === "tab") {
        inPages = true;
      }
    });
    out.push({
      style,
      indent: node.attrs.indent,
      tabStops: node.attrs.tabStops,
      text: entry,
      pages,
    });
    return true;
  });
  return out;
};

describe("insert-index command", () => {
  it("builds Index1/Index2 entries from XE fields with page numbers", () => {
    // Paragraph content starts at the paragraph pos + 1: 甲乙@1, 丙丁@6.
    const editor = build(
      docOf(body(text("甲乙"), xe("甲项")), body(text("丙丁"), xe("乙项:子项"))),
    );
    editor.commands.setTextSelection(1);
    const pages = new Map([
      [1, 2],
      [6, 5],
    ]);
    expect(editor.commands["insert-index"]((pos) => pages.get(pos) ?? 1)).toBe(true);
    const entries = entriesOf(editor);
    // `乙项:子项` nests: its main line has no page (no XE of its own — Word's
    // shape), the sub line carries it.
    expect(entries).toHaveLength(3);
    expect(entries[0]).toMatchObject({ style: "Index1", text: "甲项", pages: "2" });
    expect(entries[1]).toMatchObject({ style: "Index1", text: "乙项", pages: "" });
    // Word's built-in index indent: 220 twips per level past the first.
    expect(entries[2]).toMatchObject({
      style: "Index2",
      text: "子项",
      indent: { left: 220 },
      pages: "5",
    });
    // Every entry carries the dotted right tab stop.
    for (const entry of entries) {
      expect(entry.tabStops).toEqual([{ type: "right", position: 9350, leader: "dot" }]);
    }
    // The Index style definitions join the document styles.
    const styles = editor.state.doc.attrs.styles as {
      paragraphStyles?: { id?: string }[];
    };
    const ids = (styles.paragraphStyles ?? []).map((style) => style.id);
    expect(ids).toContain("Index1");
    expect(ids).toContain("Index2");
    editor.destroy();
  });

  it("sorts mains by the document collation and dedupes pages", () => {
    const editor = build(
      docOf(
        body(text("一"), xe("乙项"), xe("乙项")),
        body(text("二"), xe("甲项")),
        body(text("三"), xe("丙项")),
      ),
    );
    editor.commands.setTextSelection(1);
    // Pinyin collation: 丙(bing) < 甲(jia) < 乙(yi); the double 乙项 marker on
    // one page lands once.
    expect(editor.commands["insert-index"](() => 1)).toBe(true);
    expect(entriesOf(editor).map((entry) => entry.text)).toEqual(["丙项", "甲项", "乙项"]);
    expect(entriesOf(editor)[2]).toMatchObject({ pages: "1" });
    editor.destroy();
  });

  it("fails with no XE fields and inserts nothing", () => {
    const editor = build(docOf(body(text("正文"))));
    editor.commands.setTextSelection(1);
    expect(editor.commands["insert-index"]()).toBe(false);
    expect(editor.state.doc.childCount).toBe(1);
    editor.destroy();
  });
});

describe("update-index command", () => {
  it("rebuilds the entry block in place from the current XE fields", () => {
    const editor = build(docOf(body(text("甲乙"), xe("甲项")), body(text("丙丁"), xe("乙项"))));
    editor.commands.setTextSelection(1);
    editor.commands["insert-index"]();
    expect(entriesOf(editor)).toHaveLength(2);
    // Rename the XE inside the first body paragraph, then rebuild: the entry
    // follows and the stale Index paragraphs are gone (two remain, not four).
    editor.state.doc.descendants((node, at) => {
      if (node.type.name !== "inlinePassthrough") return true;
      const data = JSON.parse(String(node.attrs.data)) as {
        simpleField?: { instruction?: string };
      };
      if (data.simpleField?.instruction !== 'XE "甲项"') return true;
      editor.commands.command(({ state, dispatch }) => {
        const tr = state.tr.setNodeMarkup(at, undefined, {
          data: JSON.stringify({ simpleField: { instruction: 'XE "丙项"' } }),
        });
        dispatch?.(tr);
        return true;
      });
      return false;
    });
    expect(editor.commands["update-index"]()).toBe(true);
    expect(entriesOf(editor).map((entry) => entry.text)).toEqual(["丙项", "乙项"]);
    editor.destroy();
  });

  it("reports false when the doc has no index block", () => {
    const editor = build(docOf(body(text("甲乙"), xe("甲项"))));
    expect(editor.commands["update-index"]()).toBe(false);
    editor.destroy();
  });
});
