import { Document, InlinePassthrough, Link, Paragraph, Tab, TocField } from "@docen/docx";
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
    extensions: [Document, Paragraph, Text, Tab, Link, TocField, InlinePassthrough, TocCommands],
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

// ── Table of Figures (the \c caption switch) ──

const caption = (label: string, text: string): Record<string, unknown> => ({
  type: "paragraph",
  attrs: { style: "Caption" },
  content: [
    { type: "text", text: `${label} ` },
    {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `SEQ ${label} * ARABIC`, cachedValue: "1" },
        }),
      },
    },
    { type: "text", text: `: ${text}` },
  ],
});

describe("table-of-figures command", () => {
  it("builds one entry per matching caption with the c switch stamped", () => {
    const editor = build(
      docOf(heading(1, "Intro"), caption("Figure", "Alpha chart"), caption("Table", "Beta grid")),
    );
    editor.commands.setTextSelection(1);
    expect(editor.commands["table-of-figures"](() => 4, undefined, "Figure")).toBe(true);
    const entries = entriesOf(editor);
    // Only the Figure caption counts — Table captions belong to another \c.
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ style: "TOC1", text: "Figure : Alpha chart", page: "4" });
    expect(entries[0].linkHref).toBeNull();
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ captionLabel: "Figure" });
      }
      return true;
    });
    editor.destroy();
  });

  it("fails with no matching captions and inserts nothing", () => {
    const editor = build(docOf(caption("Table", "Beta grid")));
    editor.commands.setTextSelection(1);
    expect(editor.commands["table-of-figures"]()).toBe(false);
    expect(editor.state.doc.firstChild?.attrs.style).toBe("Caption");
    editor.destroy();
  });
});

describe("update-figures command", () => {
  it("rebuilds the figure table from current captions and keeps the label", () => {
    const editor = build(docOf(caption("Figure", "Alpha chart")));
    editor.commands.setTextSelection(1);
    editor.commands["table-of-figures"]();
    // Add a second caption, then update: the entry list follows.
    editor.commands.command(({ state, dispatch }) => {
      const node = state.schema.nodeFromJSON(caption("Figure", "Second chart"));
      dispatch?.(state.tr.insert(state.doc.content.size - 1, node));
      return true;
    });
    expect(editor.commands["update-figures"]()).toBe(true);
    expect(entriesOf(editor).map((e) => e.text)).toEqual([
      "Figure : Alpha chart",
      "Figure : Second chart",
    ]);
    editor.destroy();
  });

  it("reports false when the doc has no figure table", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(); // a heading TOC is not a figure table
    expect(editor.commands["update-figures"]()).toBe(false);
    editor.destroy();
  });
});
