import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DialogCommands, type DialogsHost } from "./dialogs";

const fieldAtom = (branch: object): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify(branch) },
  }) as JSONContent;

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const extrasOf = (): Record<string, unknown> => ({
  footnotes: [
    {
      id: 1,
      children: [{ style: "FootnoteText", children: [{ footnoteRef: true }, { text: "原始" }] }],
    },
  ],
});

const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      attrs: { documentExtras: extrasOf() },
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "ab" },
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ footnoteReference: 1 }) },
            },
            { type: "text", text: "c" },
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ endnoteReference: { id: 7 } }) },
            },
          ],
        },
      ],
    },
  });

const host = (editor: EditorType): DialogsHost => ({
  editor: () => editor,
  bridge: () => undefined,
  // No shadow root — #noteDialog()'s querySelector stays undefined and the
  // dialog opens are skipped (the commits under test don't need the element).
  element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
  syncStatusLanguage: () => {},
});

const noteOk = (detail: unknown): Event => ({ detail }) as unknown as Event;

describe("DialogCommands note targets", () => {
  it("resolves the reference atom after the caret and just before it", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    // layout: 1 "ab" | 3 footnote atom | 4 "c" | 5 endnote atom
    editor.commands.setTextSelection(4);
    expect(dialogs.noteTarget()).toEqual({ kind: "footnote", id: 1, pos: 3 });
  });

  it("reads the object-form id and the endnote kind, hit from both sides", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(5); // the atom starts here
    expect(dialogs.noteTarget()).toEqual({ kind: "endnote", id: 7, pos: 5 });
    editor.commands.setTextSelection(6); // past it — the atom ends just before
    expect(dialogs.noteTarget()).toEqual({ kind: "endnote", id: 7, pos: 5 });
  });

  it("returns null when the caret sits on plain text", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    expect(dialogs.noteTarget()).toBeNull();
  });
});

describe("DialogCommands note commits", () => {
  it("insert appends a marker-led body per textarea line", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "第一行\n第二行" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1, 2]);
    expect(notes[1].children).toEqual([
      { style: "FootnoteText", children: [{ footnoteRef: true }, { text: "第一行" }] },
      { style: "FootnoteText", children: [{ text: "第二行" }] },
    ]);
  });

  it("edit rewrites the referenced body in place, keeping style and marker", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.noteEditAtSelection(); // sets the pending target (no dialog here)
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "改后" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1]);
    expect(notes[0].children).toEqual([
      { style: "FootnoteText", children: [{ footnoteRef: true }, { text: "改后" }] },
    ]);
  });

  it("insert with empty text is a no-op", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1]);
  });
});

// Field tests: layout 1 "ab" | 3 NUMCHARS atom | 4 "c".
const buildField = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "ab" },
            fieldAtom({ simpleField: { instruction: "NUMCHARS", cachedValue: "9" } }),
            { type: "text", text: "c" },
          ],
        },
      ],
    },
  });

describe("DialogCommands field targets", () => {
  it("resolves the field atom from both sides of the caret", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4); // right after the atom
    const hit = dialogs.fieldTarget();
    expect(hit?.pos).toBe(3);
    expect(hit?.ref).toEqual({
      kind: "simpleField",
      instruction: "NUMCHARS",
      result: "9",
    });
    editor.commands.setTextSelection(3); // right at the atom
    expect(dialogs.fieldTarget()?.pos).toBe(3);
    editor.commands.setTextSelection(2); // plain text — nothing
    expect(dialogs.fieldTarget()).toBeNull();
  });
});

describe("DialogCommands field commits", () => {
  it("update re-derives the cached value from the live text", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.fieldUpdateAtSelection();
    const node = editor.state.doc.nodeAt(3);
    expect(JSON.parse(String(node?.attrs.data))).toEqual({
      simpleField: { instruction: "NUMCHARS", cachedValue: "3" },
    });
  });

  it("dialog OK on a fresh dialog seeds an evaluated field at the caret", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.fieldInsert(); // clears the (unset) edit target — insert mode
    dialogs.onFieldOk({ detail: { instruction: 'DATE \\@ "yyyy"' } } as unknown as Event);
    const seeded = JSON.parse(String(editor.state.doc.nodeAt(2)?.attrs.data));
    expect(seeded.simpleField.instruction).toBe('DATE \\@ "yyyy"');
    expect(seeded.simpleField.cachedValue).toBe("2026");
  });

  it("dialog OK after edit rewrites the atom's instruction in place", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.fieldEditAtSelection();
    dialogs.onFieldOk({ detail: { instruction: 'DATE \\@ "yy"' } } as unknown as Event);
    const node = editor.state.doc.nodeAt(3);
    expect(JSON.parse(String(node?.attrs.data))).toEqual({
      simpleField: { instruction: 'DATE \\@ "yy"', cachedValue: "26" },
    });
  });

  it("checkbox toggle flips checked and keeps the rest of the branch", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [fieldAtom({ formField: { checkBox: { checked: false, size: 20 } } })],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(1);
    expect(dialogs.fieldTarget()?.ref.kind).toBe("formField");
    dialogs.fieldToggleCheckboxAtSelection();
    expect(JSON.parse(String(editor.state.doc.nodeAt(1)?.attrs.data))).toEqual({
      formField: { checkBox: { checked: true, size: 20 } },
    });
  });
});
