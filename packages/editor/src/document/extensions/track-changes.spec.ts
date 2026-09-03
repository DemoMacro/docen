import { Deletion, Document, Insertion, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { TrackChanges } from "./track-changes";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const docOf = (text: string): Record<string, unknown> => ({
  type: "doc",
  content: [{ type: "paragraph", content: [{ type: "text", text }] }],
});

/** A headless editor with exactly the schema the track-changes workflow
 *  touches: paragraph text + the two revision marks + the extension. */
const build = (initial = "hello world"): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Insertion, Deletion, TrackChanges],
    content: docOf(initial),
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

/** Raw PM edit through the chain — the same path the canvas input bridge and
 *  every Tiptap command take. */
const type = (editor: Editor, text: string, from?: number, to?: number): void => {
  editor.commands.command(({ state, dispatch }) => {
    const start = from ?? state.selection.from;
    const end = to ?? state.selection.to;
    dispatch?.(state.tr.insertText(text, start, end));
    return true;
  });
};

const marksOf = (editor: Editor): { type: string; attrs: Record<string, unknown> }[] => {
  const out: { type: string; attrs: Record<string, unknown> }[] = [];
  editor.state.doc.descendants((node) => {
    if (!node.isText) return true;
    for (const mark of node.marks) {
      if (mark.type.name === "insertion" || mark.type.name === "deletion") {
        out.push({ type: mark.type.name, attrs: mark.attrs as Record<string, unknown> });
      }
    }
    return true;
  });
  return out;
};

const textOf = (editor: Editor): string =>
  editor.state.doc.textBetween(0, editor.state.doc.content.size, "\n");

describe("TrackChanges marking", () => {
  it("off by default — edits apply bare", () => {
    const editor = build();
    type(editor, "XY", 6, 6);
    expect(textOf(editor)).toBe("helloXY world");
    expect(marksOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("marks typed text as an insertion record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "XY", 6, 6);
    const marks = marksOf(editor);
    expect(marks).toHaveLength(1);
    expect(marks[0].type).toBe("insertion");
    expect(marks[0].attrs.author).toBe("docen");
    expect(marks[0].attrs.id).toBe(1);
    expect(String(marks[0].attrs.date)).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
    editor.destroy();
  });

  it("merges consecutive same-author typing into one record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "X", 6, 6);
    type(editor, "Y", 7, 7);
    expect(marksOf(editor)).toHaveLength(1);
    editor.destroy();
  });

  it("keeps deleted text under a deletion record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 6, 9); // strike " wo"
    // Struck text stays in the flow — the visible string is unchanged.
    expect(textOf(editor)).toBe("hello world");
    expect(marksOf(editor)).toEqual([
      { type: "deletion", attrs: expect.objectContaining({ author: "docen" }) },
    ]);
    editor.destroy();
  });

  it("refuses to delete already-struck text — it stays struck", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 6, 9);
    const struck = textOf(editor);
    type(editor, "", 6, 9); // the struck range occupies [6,9) again
    expect(textOf(editor)).toBe(struck);
    expect(marksOf(editor)).toHaveLength(1);
    editor.destroy();
  });

  it("keeps the caret before struck text so consecutive Backspace marks on", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 11, 12); // backspace over "d"
    // The caret must land BEFORE the restored "d" — the canvas bridge derives
    // the next Backspace's range from it, and a caret behind the struck runs
    // retargets the same character forever (stalled consecutive deletes).
    expect(editor.state.selection.from).toBe(11);
    type(editor, "", editor.state.selection.from - 1, editor.state.selection.from); // "l"
    expect(textOf(editor)).toBe("hello world");
    expect(editor.state.selection.from).toBe(10);
    // "l" and "d" merge into ONE struck run under ONE record (PM normalizes
    // the adjacent equal-mark texts) — Word's consecutive-delete semantics.
    expect(marksOf(editor)).toHaveLength(1);
    expect(marksOf(editor)[0]!.type).toBe("deletion");
    editor.destroy();
  });

  it("crosses already-struck text instead of stalling on it", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 6, 9); // strike " wo"
    expect(editor.state.selection.from).toBe(6);
    type(editor, "", 8, 9); // a delete that hits the struck "o" — refused
    expect(textOf(editor)).toBe("hello world");
    expect(marksOf(editor)).toHaveLength(1);
    // But the caret crosses it (Word), so editing continues further left.
    expect(editor.state.selection.from).toBe(8);
    editor.destroy();
  });

  it("replacing a selection keeps struck original after the inserted text", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "Z", 1, 6); // replace "hello"
    expect(textOf(editor)).toBe("Zhello world");
    expect(marksOf(editor).map((m) => m.type)).toEqual(["insertion", "deletion"]);
    editor.destroy();
  });

  it("leaves structural edits untracked", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.split(state.selection.from));
      return true;
    });
    expect(marksOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("skips undo-shaped transactions (history$ meta)", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "X", 6, 6);
    editor.commands.command(({ state, dispatch }) => {
      const tr = state.tr;
      // prosemirror-history tags its replay trs with the "history$" meta;
      // the marking plugin must let them through untouched.
      tr.setMeta("history$", { redo: false });
      dispatch?.(tr.insertText("Y", 6, 6));
      return true;
    });
    expect(marksOf(editor)).toHaveLength(1); // only the first insertion
    editor.destroy();
  });
});

describe("TrackChanges accept/reject", () => {
  it("accept removes the insertion mark and keeps the text", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "XY", 6, 6);
    editor.commands["accept-change"]();
    expect(textOf(editor)).toBe("helloXY world");
    expect(marksOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("accept deletes struck text", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 6, 9); // strike " wo"
    editor.commands["accept-change"]();
    // The struck " wo" leaves for real — "hello" and "rld" join up.
    expect(textOf(editor)).toBe("hellorld");
    expect(marksOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("reject restores struck text", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "", 6, 9);
    editor.commands["reject-change"]();
    expect(textOf(editor)).toBe("hello world");
    expect(marksOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("reject removes inserted text", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    type(editor, "XY", 6, 6);
    editor.commands["reject-change"]();
    expect(textOf(editor)).toBe("hello world");
    editor.destroy();
  });

  it("navigates revisions with next/previous", () => {
    const editor = build("aaaa bbbb cccc");
    editor.commands["track-changes"](true);
    // Seed two records without the plugin: "aaaa" and "bbbb".
    editor.commands.command(({ state, dispatch }) => {
      const insertion = state.schema.marks.insertion!;
      const attrs = { id: 1, author: "docen", date: "2026-01-01T00:00:00Z" };
      dispatch?.(
        state.tr
          .addMark(1, 5, insertion.create(attrs))
          .addMark(6, 10, insertion.create({ ...attrs, id: 2 })),
      );
      return true;
    });
    editor.commands["next-change"](); // caret starts at 1 → first range after
    expect(editor.state.selection.from).toBe(6);
    editor.commands["previous-change"]();
    expect(editor.state.selection.from).toBe(1);
    editor.destroy();
  });

  it("accept/reject report false with no revision present", () => {
    const editor = build();
    expect(editor.commands["accept-change"]()).toBe(false);
    expect(editor.commands["reject-change"]()).toBe(false);
    editor.destroy();
  });
});
