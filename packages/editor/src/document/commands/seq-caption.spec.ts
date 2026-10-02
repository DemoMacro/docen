import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import type { FieldFrame } from "../fields";
import { DialogCommands, type DialogsHost } from "./dialogs";

// Caption & SEQ auto-numbering: the dialog's inserted field, its renumbering
// through Update All Fields (delete/reorder), chapter numbering (`\s`), the
// `\*` number formats, and the insert position. The evaluator oracle values
// are Word's documented results (Figure 1, Figure 1-2, Figure I, …).

const fieldAtom = (branch: object): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify(branch) },
  }) as JSONContent;

const field = (instruction: string, cachedValue: string): JSONContent =>
  fieldAtom({ simpleField: { instruction, cachedValue } });

const Text = TextNode.create({ name: "text", group: "inline" });

const heading = (level: number, text: string): JSONContent => ({
  type: "paragraph",
  attrs: { heading: `Heading${level}` },
  content: [{ type: "text", text }],
});

const paragraph = (text: string): JSONContent => ({
  type: "paragraph",
  content: [{ type: "text", text }],
});

const seqParagraph = (instruction: string, cached = "9"): JSONContent => ({
  type: "paragraph",
  content: [field(instruction, cached)],
});

const build = (content: JSONContent[], attrs?: Record<string, unknown>): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: { type: "doc", ...(attrs ? { attrs } : {}), content },
  });

/** The host DialogCommands runs against — the caption commit path only needs
 *  the editor and the (absent) bridge. */
const host = (
  editor: EditorType,
  frame?: (pos: number) => FieldFrame | undefined,
): DialogsHost => ({
  editor: () => editor,
  bridge: () => ({ activeEditor: () => editor, focus: () => {}, pageOf: () => null }),
  element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
  syncStatusLanguage: () => {},
  fieldFrame: frame,
});

/** Put the caret inside the paragraph whose text is exactly `text`. */
const caretIn = (editor: EditorType, text: string): void => {
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === "paragraph" && node.textContent === text) {
      editor.commands.setTextSelection(pos + 1);
      return false;
    }
    return true;
  });
};

/** Every simple field in document order, with the atom's position. */
const fieldsOf = (editor: EditorType): { pos: number; instruction: string; cached: string }[] => {
  const out: { pos: number; instruction: string; cached: string }[] = [];
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== "inlinePassthrough") return true;
    try {
      const data = JSON.parse(String(node.attrs.data ?? "{}")) as {
        simpleField?: { instruction?: string; cachedValue?: string };
      };
      if (data.simpleField)
        out.push({
          pos,
          instruction: data.simpleField.instruction ?? "",
          cached: data.simpleField.cachedValue ?? "",
        });
    } catch {
      /* opaque blob */
    }
    return true;
  });
  return out;
};

const captionsOf = (editor: EditorType): { style: unknown; text: string }[] => {
  const out: { style: unknown; text: string }[] = [];
  editor.state.doc.forEach((node) => {
    if (node.attrs.style === "Caption")
      out.push({ style: node.attrs.style, text: node.textContent });
  });
  return out;
};

const blockTextsOf = (editor: EditorType): string[] => {
  const out: string[] = [];
  editor.state.doc.forEach((node) =>
    out.push(node.attrs.style === "Caption" ? "caption" : node.textContent),
  );
  return out;
};

const captionSettingsOf = (editor: EditorType): Record<string, unknown>[] => {
  const extras = editor.state.doc.attrs.documentExtras as {
    settings?: { captions?: { captions?: Record<string, unknown>[] } };
  };
  return extras?.settings?.captions?.captions ?? [];
};

/** Delete the first paragraph whose text matches `text`. */
const deleteParagraph = (editor: EditorType, text: string): void => {
  editor.commands.command(({ state, dispatch }) => {
    let range: [number, number] | null = null;
    state.doc.forEach((node, offset) => {
      if (node.textContent === text) range = [offset, offset + node.nodeSize];
    });
    if (range) dispatch?.(state.tr.delete(range[0], range[1]));
    return true;
  });
};

describe("caption commit", () => {
  it("seeds an evaluated SEQ field and renumbers the next insert", () => {
    const editor = build([paragraph("内容"), paragraph("尾部")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Alpha chart", position: "below" },
    } as unknown as Event);
    let fields = fieldsOf(editor);
    expect(fields).toHaveLength(1);
    expect(fields[0]!.instruction).toBe("SEQ Figure \\* ARABIC");
    expect(fields[0]!.cached).toBe("1");
    expect(captionsOf(editor)).toEqual([{ style: "Caption", text: "Figure : Alpha chart" }]);

    // A second caption continues the label's sequence from the document.
    caretIn(editor, "尾部");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Beta chart", position: "below" },
    } as unknown as Event);
    fields = fieldsOf(editor);
    expect(fields.map((f) => f.cached)).toEqual(["1", "2"]);
    expect(captionsOf(editor).map((c) => c.text)).toEqual([
      "Figure : Alpha chart",
      "Figure : Beta chart",
    ]);
    editor.destroy();
  });

  it("inserts above the anchored paragraph when asked", () => {
    const editor = build([paragraph("内容"), paragraph("尾部")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Alpha chart", position: "above" },
    } as unknown as Event);
    expect(captionsOf(editor)).toEqual([{ style: "Caption", text: "Figure : Alpha chart" }]);
    // The caption precedes the anchored paragraph; the body follows it.
    expect(blockTextsOf(editor)).toEqual(["caption", "内容", "尾部"]);
    editor.destroy();
  });

  it("carries the format switch and chapter shape into the field code and settings", () => {
    const editor = build([heading(1, "第一章"), paragraph("内容")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: {
        label: "Figure",
        text: "Alpha chart",
        position: "below",
        chapterNumber: true,
        heading: 1,
        sep: "colon",
        format: "ROMAN",
      },
    } as unknown as Event);
    const fields = fieldsOf(editor);
    expect(fields).toHaveLength(1);
    expect(fields[0]!.instruction).toBe("SEQ Figure \\* ROMAN \\s 1");
    // Chapter 1 + colon + the Roman sequence number.
    expect(fields[0]!.cached).toBe("1:I");
    expect(captionSettingsOf(editor)).toEqual([
      { name: "Figure", chapterNumber: true, heading: 1, sep: "colon", numFmt: "upperRoman" },
    ]);
    // The persisted separator drives the update; the value is stable.
    expect(dialogs.updateAllFields()).toBe(0);
    editor.destroy();
  });

  it("keeps other labels' settings when a second label is inserted", () => {
    const editor = build([paragraph("内容")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "A", position: "below", chapterNumber: true, heading: 1 },
    } as unknown as Event);
    dialogs.onCaptionOk({
      detail: { label: "Table", text: "B", position: "below", sep: "period" },
    } as unknown as Event);
    expect(captionSettingsOf(editor)).toEqual([
      { name: "Figure", chapterNumber: true, heading: 1, sep: "hyphen", numFmt: "decimal" },
      { name: "Table", chapterNumber: false, sep: "period", numFmt: "decimal" },
    ]);
    editor.destroy();
  });

  it("starts a new chapter's first caption at 1 (pending \\s reset at the insert point)", () => {
    const editor = build([
      heading(1, "第一章"),
      paragraph("第一节内容"),
      heading(1, "第二章"),
      paragraph("第二节内容"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "第一节内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "A", position: "below", chapterNumber: true, heading: 1 },
    } as unknown as Event);
    caretIn(editor, "第二节内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "B", position: "below", chapterNumber: true, heading: 1 },
    } as unknown as Event);
    // The second insert must see the chapter-2 heading's reset even though no
    // SEQ atom sits between the heading and the insertion point.
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1-1", "2-1"]);
    editor.destroy();
  });
});

describe("updateAllFields SEQ numbering", () => {
  it("renumbers captions in document order after one is deleted", () => {
    const editor = build([
      seqParagraph("SEQ Figure \\* ARABIC"),
      seqParagraph("SEQ Figure \\* ARABIC"),
      seqParagraph("SEQ Figure \\* ARABIC"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(3);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1", "2", "3"]);

    // Delete the first caption: the remaining two close the gap.
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.delete(0, state.doc.firstChild!.nodeSize));
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(2);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1", "2"]);
    editor.destroy();
  });

  it("follows a moved caption's new document position", () => {
    const editor = build([
      {
        type: "paragraph",
        content: [{ type: "text", text: "A " }, field("SEQ Figure \\* ARABIC", "1")],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "B " }, field("SEQ Figure \\* ARABIC", "2")],
      },
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(0); // already in order

    // Move the first caption block after the second: the counters swap.
    editor.commands.command(({ state, dispatch }) => {
      const first = state.doc.firstChild!;
      const tr = state.tr.delete(0, first.nodeSize);
      tr.insert(tr.doc.content.size, first);
      dispatch?.(tr);
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(2);
    expect(blockTextsOf(editor)).toEqual(["B ", "A "]);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1", "2"]);
    editor.destroy();
  });

  it("renumbers caption fields the dialog inserted after a delete", () => {
    const editor = build([paragraph("内容"), paragraph("尾部")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Alpha", position: "below" },
    } as unknown as Event);
    caretIn(editor, "尾部");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Beta", position: "below" },
    } as unknown as Event);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1", "2"]);

    deleteParagraph(editor, "Figure : Alpha");
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1"]);
    expect(captionsOf(editor).map((c) => c.text)).toEqual(["Figure : Beta"]);
    editor.destroy();
  });

  it("updates the field under the caret (F9) from its live SEQ position", () => {
    const editor = build([
      heading(1, "第一章"),
      {
        type: "paragraph",
        content: [{ type: "text", text: "x" }, field("SEQ Figure \\* ARABIC", "9")],
      },
    ]);
    const dialogs = new DialogCommands(host(editor));
    const atomPos = fieldsOf(editor)[0]!.pos;
    editor.commands.setTextSelection(atomPos + 1);
    dialogs.fieldUpdateAtSelection();
    expect(fieldsOf(editor)[0]!.cached).toBe("1");
    editor.destroy();
  });

  it("F9 restarts a stale caption at its chapter boundary", () => {
    const editor = build([
      heading(1, "第一章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1", "1-1"),
      heading(1, "第二章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1", "1-9"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    const atomPos = fieldsOf(editor)[1]!.pos;
    editor.commands.setTextSelection(atomPos + 1);
    dialogs.fieldUpdateAtSelection();
    // The caret-bound walk applies the chapter heading's pending reset: the
    // caption is the first of chapter 2, not the second of chapter 1.
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1-1", "2-1"]);
    editor.destroy();
  });

  it("insert-field seeds the chapter-prefixed value in a new chapter", () => {
    const editor = build([
      heading(1, "第一章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1", "1-1"),
      heading(1, "第二章"),
      paragraph("x"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "x");
    dialogs.onFieldOk({
      detail: { instruction: "SEQ Figure \\* ARABIC \\s 1" },
    } as unknown as Event);
    const fields = fieldsOf(editor);
    expect(fields).toHaveLength(2);
    expect(fields[1]!.cached).toBe("2-1");
    editor.destroy();
  });
});

describe("chapter numbering (SEQ \\s)", () => {
  it("prefixes the chapter number and restarts at each heading level", () => {
    const editor = build([
      heading(1, "第一章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1"),
      heading(2, "小节"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1"),
      heading(1, "第二章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 1"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(3);
    // Chapter 1: two figures; chapter 2 restarts at one. A heading 2 does not
    // reset a level-1 sequence.
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1-1", "1-2", "2-1"]);
    editor.destroy();
  });

  it("uses the persisted caption separator, hyphen by default", () => {
    const content = [heading(1, "第一章"), seqParagraph("SEQ Figure \\* ARABIC \\s 1")];
    const withPeriod = build(content, {
      documentExtras: {
        settings: {
          captions: {
            captions: [{ name: "Figure", chapterNumber: true, heading: 1, sep: "period" }],
          },
        },
      },
    });
    expect(new DialogCommands(host(withPeriod)).updateAllFields()).toBe(1);
    expect(fieldsOf(withPeriod).map((f) => f.cached)).toEqual(["1.1"]);

    const plain = build(content);
    expect(new DialogCommands(host(plain)).updateAllFields()).toBe(1);
    expect(fieldsOf(plain).map((f) => f.cached)).toEqual(["1-1"]);
    withPeriod.destroy();
    plain.destroy();
  });

  it("keeps the plain sequence number when no heading of the level precedes", () => {
    const editor = build([seqParagraph("SEQ Figure \\* ARABIC \\s 2")]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1"]);
    editor.destroy();
  });

  it("restarts the sequence when a heading above the \\s level starts a new chapter", () => {
    const editor = build([
      heading(2, "节一"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 2"),
      heading(1, "第一章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 2"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(2);
    // "节一" opens chapter 1 (Heading 2); the Heading 1 starts a new chapter
    // and restarts the sequence. The second caption has no chapter prefix —
    // no Heading 2 precedes it in the new chapter (Word's no-chapter shape).
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1-1", "1"]);
    editor.destroy();
  });

  it("ignores a non-integer \\s level (no reset, no chapter prefix)", () => {
    const editor = build([
      heading(1, "第一章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 2.5"),
      heading(1, "第二章"),
      seqParagraph("SEQ Figure \\* ARABIC \\s 2.5"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(2);
    // Word ignores the malformed switch: the sequence continues across the
    // heading and no chapter prefix appears.
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["1", "2"]);
    editor.destroy();
  });
});

describe("format switches (SEQ \\*)", () => {
  it("formats ARABIC, ROMAN/roman, ALPHABETIC/alphabetic per label", () => {
    const editor = build([
      {
        type: "paragraph",
        content: [field("SEQ Roman \\* ROMAN", "9"), field("SEQ Roman \\* ROMAN", "9")],
      },
      { type: "paragraph", content: [field("SEQ Romanic \\* roman", "9")] },
      {
        type: "paragraph",
        content: [field("SEQ Alpha \\* ALPHABETIC", "9"), field("SEQ Alpha \\* alphabetic", "9")],
      },
      {
        type: "paragraph",
        content: [field("SEQ Arab \\* ARABIC", "9"), field("SEQ Arab", "9")],
      },
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(7);
    // One shared counter per label; the switch picks the glyphs.
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["I", "II", "i", "A", "b", "1", "2"]);
    editor.destroy();
  });

  it("reads the \\* switch case-insensitively; the casing picks the output case", () => {
    const editor = build([
      {
        type: "paragraph",
        content: [
          field("SEQ Roman \\* Roman", "9"),
          field("SEQ Roman \\* ROMAN", "9"),
          field("SEQ Alpha \\* Alphabetic", "9"),
          field("SEQ Alpha \\* ALPHABETIC", "9"),
          field("SEQ Arabic \\* Arabic", "9"),
          field("SEQ Arabic \\* arabic", "9"),
        ],
      },
    ]);
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(6);
    expect(fieldsOf(editor).map((f) => f.cached)).toEqual(["I", "II", "A", "B", "1", "2"]);
    editor.destroy();
  });
});
