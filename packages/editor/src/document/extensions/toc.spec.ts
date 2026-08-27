import { Document, Link, Paragraph, Tab, TocField } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { TocCommands } from "./toc";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const heading = (level: number, text: string): Record<string, unknown> => ({
  type: "paragraph",
  attrs: { heading: `Heading${level}` },
  content: [{ type: "text", text }],
});

const docOf = (...blocks: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "doc",
  content: blocks,
});

/** A headless editor with exactly the schema the TOC workflow touches. */
const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Tab, Link, TocField, TocCommands],
    content: doc,
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

interface EntryInfo {
  style: unknown;
  indent: unknown;
  tabStops: unknown;
  text: string;
  linkHref: unknown;
  page: string;
}

/** Flatten the first tocField's entries into testable shapes. */
const entriesOf = (editor: EditorType): EntryInfo[] => {
  const out: EntryInfo[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "tocField") return true;
    node.forEach((entry) => {
      let text = "";
      let linkHref: unknown = null;
      let page = "";
      let inPage = false;
      entry.content.forEach((child) => {
        if (child.type.name === "text") {
          if (inPage) page += child.text ?? "";
          else text += child.text ?? "";
          const link = child.marks.find((m) => m.type.name === "link");
          if (link) linkHref = link.attrs.href;
        } else if (child.type.name === "tab") {
          inPage = true;
        }
      });
      out.push({
        style: entry.attrs.style,
        indent: entry.attrs.indent,
        tabStops: entry.attrs.tabStops,
        text,
        linkHref,
        page,
      });
    });
    return false;
  });
  return out;
};

describe("toc command", () => {
  it("builds one entry per heading 1-3 with style, tab, link, and page", () => {
    const editor = build(
      docOf(heading(1, "Alpha"), heading(2, "Beta"), heading(3, "Gamma"), {
        type: "paragraph",
        content: [{ type: "text", text: "body" }],
      }),
    );
    editor.commands.setTextSelection(1);
    // Doc layout (paragraph content starts at the paragraph pos + 1):
    // Alpha@1, Beta@8, Gamma@14.
    const pages = new Map([
      [1, 2],
      [8, 5],
      [14, 9],
    ]);
    const pageOf = (pos: number): number => pages.get(pos) ?? 1;
    expect(editor.commands.toc(pageOf)).toBe(true);
    const entries = entriesOf(editor);
    expect(entries).toHaveLength(3);
    expect(entries[0]).toMatchObject({
      style: "TOC1",
      text: "Alpha",
      linkHref: "#_Toc1",
      page: "2",
    });
    // Word's built-in TOC indent: 220 twips per level past the first.
    expect(entries[1]).toMatchObject({
      style: "TOC2",
      text: "Beta",
      indent: { left: 220 },
      page: "5",
    });
    expect(entries[2]).toMatchObject({ style: "TOC3", text: "Gamma", indent: { left: 440 } });
    // Every entry carries the dotted right tab stop.
    for (const entry of entries) {
      expect(entry.tabStops).toEqual([{ type: "right", position: 9350, leader: "dot" }]);
    }
    editor.destroy();
  });

  it("skips headings past level 3 and empty headings", () => {
    const editor = build(
      docOf(heading(1, "One"), heading(4, "Deep"), {
        type: "paragraph",
        attrs: { heading: "Heading2" },
      }),
    );
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    expect(entriesOf(editor)).toHaveLength(1);
    editor.destroy();
  });

  it("fails with no headings and inserts nothing", () => {
    const editor = build(docOf({ type: "paragraph" }));
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc()).toBe(false);
    expect(editor.state.doc.firstChild?.type.name).toBe("paragraph");
    editor.destroy();
  });

  it("stamps the default field switches on the tocField", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ headingStyleRange: "1-3", hyperlink: true });
      }
      return true;
    });
    editor.destroy();
  });
});

describe("update-toc command", () => {
  it("rebuilds entries from current headings and keeps the field switches", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(2, "Beta")));
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    // Rename a heading, then update: the entry text follows. The inserted TOC
    // sits before the headings, so "Alpha" now lives at [18, 23).
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("Renamed", 18, 23)); // replace "Alpha"
      return true;
    });
    expect(editor.commands["update-toc"]()).toBe(true);
    const entries = entriesOf(editor);
    expect(entries.map((e) => e.text)).toEqual(["Renamed", "Beta"]);
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ headingStyleRange: "1-3", hyperlink: true });
      }
      return true;
    });
    editor.destroy();
  });

  it("reports false when the doc has no tocField", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    expect(editor.commands["update-toc"]()).toBe(false);
    editor.destroy();
  });
});
