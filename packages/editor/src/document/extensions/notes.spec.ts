import { Document, InlinePassthrough, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it } from "vitest";

import { liveNoteIds, noteRefId, NotesCleanup, prunedExtras } from "./notes";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const refAtom = (branch: object): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify(branch) },
});

const build = (extras?: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough, UndoRedo, NotesCleanup],
    content: {
      type: "doc",
      ...(extras ? { attrs: { documentExtras: extras } } : {}),
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "a" },
            refAtom({ footnoteReference: 1 }),
            { type: "text", text: "b" },
            refAtom({ footnoteReference: { id: 2 } }),
            { type: "text", text: "c" },
            refAtom({ endnoteReference: 3 }),
            refAtom({ bookmarkStart: { name: "B" } }),
          ],
        },
      ],
    },
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

const noteIdsIn = (
  extras: unknown,
): { footnotes: Array<number | undefined>; endnotes: Array<number | undefined> } => {
  const e = (extras ?? {}) as {
    footnotes?: Array<{ id?: number }>;
    endnotes?: Array<{ id?: number }>;
  };
  return {
    footnotes: (e.footnotes ?? []).map((n) => n.id),
    endnotes: (e.endnotes ?? []).map((n) => n.id),
  };
};

const deleteAt = (editor: EditorType, from: number, to: number): void => {
  editor.commands.command(({ state, dispatch }) => {
    dispatch?.(state.tr.delete(from, to));
    return true;
  });
};

const atomPos = (editor: EditorType, branch: object): number => {
  const data = JSON.stringify(branch);
  let pos = -1;
  editor.state.doc.descendants((node, p) => {
    if (pos < 0 && node.type.name === "inlinePassthrough" && node.attrs.data === data) pos = p;
    return true;
  });
  return pos;
};

describe("noteRefId", () => {
  it("reads the flat number form and the option-object form", () => {
    expect(noteRefId({ footnoteReference: 1 })).toEqual({ kind: "footnoteReference", id: 1 });
    expect(noteRefId({ footnoteReference: { id: 2 } })).toEqual({
      kind: "footnoteReference",
      id: 2,
    });
    expect(noteRefId({ endnoteReference: 3 })).toEqual({ kind: "endnoteReference", id: 3 });
  });

  it("returns null for non-reference branches", () => {
    expect(noteRefId({ bookmarkStart: { name: "B" } })).toBeNull();
    expect(noteRefId({ footnoteReference: { noId: true } })).toBeNull();
    expect(noteRefId({})).toBeNull();
  });
});

describe("liveNoteIds", () => {
  it("collects every referenced id per channel from the passthrough atoms", () => {
    const editor = build();
    expect(liveNoteIds(editor.state.doc)).toEqual({
      footnoteReference: new Set([1, 2]),
      endnoteReference: new Set([3]),
    });
  });
});

describe("prunedExtras", () => {
  it("drops orphaned entries, keeps live and id-less ones, passes other keys through", () => {
    const pruned = prunedExtras(
      {
        footnotes: [{ id: 1 }, { id: 2 }, { children: [] }],
        endnotes: [{ id: 3 }],
        sectionProperties: { kept: true },
      },
      { footnoteReference: new Set([2]), endnoteReference: new Set() },
    );
    expect(pruned).toEqual({
      footnotes: [{ id: 2 }, { children: [] }],
      endnotes: [],
      sectionProperties: { kept: true },
    });
  });

  it("prunes both channels in one pass without losing the first channel's drop", () => {
    const pruned = prunedExtras(
      { footnotes: [{ id: 1 }], endnotes: [{ id: 3 }] },
      { footnoteReference: new Set(), endnoteReference: new Set() },
    );
    expect(pruned).toEqual({ footnotes: [], endnotes: [] });
  });

  it("returns null when nothing is orphaned", () => {
    const extras = {
      footnotes: [{ id: 1 }, { id: 2 }],
      endnotes: [{ id: 3 }],
    };
    expect(
      prunedExtras(extras, {
        footnoteReference: new Set([1, 2]),
        endnoteReference: new Set([3]),
      }),
    ).toBeNull();
  });
});

describe("NotesCleanup plugin", () => {
  const extrasOf = (): Record<string, unknown> => ({
    footnotes: [
      { id: 1, children: [{ children: [{ text: "one" }] }] },
      { id: 2, children: [{ children: [{ text: "two" }] }] },
    ],
    endnotes: [{ id: 3, children: [{ children: [{ text: "three" }] }] }],
  });

  it("prunes the body when its last reference is deleted", () => {
    const editor = build(extrasOf());
    const pos = atomPos(editor, { footnoteReference: 1 });
    deleteAt(editor, pos, pos + 1);
    expect(noteIdsIn(editor.state.doc.attrs.documentExtras)).toEqual({
      footnotes: [2],
      endnotes: [3],
    });
  });

  it("one undo restores the reference and the pruned body together", () => {
    const editor = build(extrasOf());
    const pos = atomPos(editor, { endnoteReference: 3 });
    deleteAt(editor, pos, pos + 1);
    expect(noteIdsIn(editor.state.doc.attrs.documentExtras).endnotes).toEqual([]);
    editor.commands.undo();
    expect(noteIdsIn(editor.state.doc.attrs.documentExtras)).toEqual({
      footnotes: [1, 2],
      endnotes: [3],
    });
  });

  it("plain typing never touches the note arrays", () => {
    const editor = build(extrasOf());
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("x", state.selection.from));
      return true;
    });
    expect(noteIdsIn(editor.state.doc.attrs.documentExtras)).toEqual({
      footnotes: [1, 2],
      endnotes: [3],
    });
  });

  it("deleting text that orphans nothing leaves the arrays alone", () => {
    const editor = build(extrasOf());
    deleteAt(editor, 1, 2); // the leading "a"
    expect(noteIdsIn(editor.state.doc.attrs.documentExtras)).toEqual({
      footnotes: [1, 2],
      endnotes: [3],
    });
  });

  it("survives a document without documentExtras", () => {
    const editor = build();
    const pos = atomPos(editor, { footnoteReference: 1 });
    deleteAt(editor, pos, pos + 1);
    expect(editor.state.doc.attrs.documentExtras ?? null).toBeNull();
  });
});
