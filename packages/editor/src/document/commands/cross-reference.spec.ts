import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import {
  CROSS_REFERENCE_CONTENTS,
  crossReferenceContentAvailable,
  crossReferenceInstruction,
  type CrossRefContent,
  type CrossReferenceTarget,
} from "../cross-reference";
import { resolveField, type FieldBookmark, type FieldFrame } from "../fields";
import { DialogCommands, type DialogsHost } from "./dialogs";

// Cross-reference 2.0: the type × content commit matrix (heading/numbered/
// bookmark/note/caption), Word's hidden `_Ref…` bookmarks around referenced
// targets, REF/PAGEREF/NOTEREF update semantics (text, page, `\n` number,
// `\p` above/below), and caption slices with chapter-prefixed numbers.

const fieldAtom = (branch: object): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify(branch) },
  }) as JSONContent;

const Text = TextNode.create({ name: "text", group: "inline" });

const paragraph = (text: string): JSONContent => ({
  type: "paragraph",
  content: [{ type: "text", text }],
});

const heading = (level: number, text: string, attrs?: Record<string, unknown>): JSONContent => ({
  type: "paragraph",
  attrs: { heading: `Heading${level}`, ...attrs },
  content: [{ type: "text", text }],
});

/** The doc's numbering definitions for the numbered-item specs. */
const NUMBERING = {
  abstractNumberings: [
    {
      reference: "list_1",
      levels: [
        { level: 0, format: "decimal", text: "%1." },
        { level: 1, format: "lowerLetter", text: "%1.%2." },
      ],
    },
  ],
};

const build = (content: JSONContent[], attrs?: Record<string, unknown>): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: { type: "doc", ...(attrs ? { attrs } : {}), content },
  });

const host = (
  editor: EditorType,
  opts: { frame?: (pos: number) => FieldFrame | undefined; pageOf?: () => number | null } = {},
): DialogsHost => ({
  editor: () => editor,
  bridge: () => ({
    activeEditor: () => editor,
    focus: () => {},
    pageOf: () => opts.pageOf?.() ?? null,
  }),
  element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
  syncStatusLanguage: () => {},
  filename: () => "报告.docx",
  fieldFrame: opts.frame,
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

/** Bookmark pairs in document order: name and the inner range. */
const bookmarksOf = (
  editor: EditorType,
): { name: string; from: number; to: number; id: number }[] => {
  const out: { name: string; from: number; to: number; id: number }[] = [];
  const open: { name: string; from: number; id: number }[] = [];
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== "inlinePassthrough") return true;
    try {
      const data = JSON.parse(String(node.attrs.data ?? "{}")) as {
        bookmarkStart?: { id?: number; name?: string };
        bookmarkEnd?: { id?: number };
      };
      if (data.bookmarkStart?.name)
        open.push({
          name: data.bookmarkStart.name,
          from: pos + node.nodeSize,
          id: data.bookmarkStart.id ?? 0,
        });
      if (data.bookmarkEnd) {
        const at = open.findIndex((entry) => entry.id === data.bookmarkEnd!.id);
        const hit = open[at >= 0 ? at : 0];
        if (hit) {
          out.push({ ...hit, to: pos });
          open.splice(at >= 0 ? at : 0, 1);
        }
      }
    } catch {
      /* opaque blob */
    }
    return true;
  });
  return out;
};

const commit = (dialogs: DialogCommands, key: string, content: CrossRefContent): void => {
  dialogs.onCrossRefOk({ detail: { key, content } } as unknown as Event);
};

const targetOfKind = (
  dialogs: DialogCommands,
  kind: string,
  text?: string,
): ReturnType<DialogCommands["crossReferenceTargets"]>[number] => {
  const hit = dialogs
    .crossReferenceTargets()
    .find((entry) => entry.kind === kind && (text == null || entry.listText.includes(text)));
  if (!hit) throw new Error(`no ${kind} target${text ? ` matching ${text}` : ""}`);
  return hit;
};

/** A frame function keyed by the paragraph's text, so a spec can move a
 *  target to another page by mutating the map. */
const frameByText = (
  editor: EditorType,
  pages: Map<string, number>,
): ((pos: number) => FieldFrame | undefined) => {
  return (pos) => {
    let text: string | null = null;
    editor.state.doc.descendants((node, at) => {
      if (text != null) return false;
      if (node.type.name === "paragraph" && pos >= at && pos < at + node.nodeSize) {
        text = node.textContent;
        return false;
      }
      return true;
    });
    const page = text != null ? pages.get(text) : undefined;
    return page != null
      ? { page, pageCount: 9, section: 1, sectionPages: 9, pageFormat: "decimal" }
      : undefined;
  };
};

describe("cross-reference field instructions", () => {
  it("maps each type × content to Word's field shape", () => {
    expect(crossReferenceInstruction("heading", "text", "H")).toBe("REF H \\h");
    expect(crossReferenceInstruction("heading", "number", "H")).toBe("REF H \\n \\h");
    expect(crossReferenceInstruction("numbered", "number", "N")).toBe("REF N \\n \\h");
    expect(crossReferenceInstruction("bookmark", "text", "B")).toBe("REF B \\h");
    expect(crossReferenceInstruction("bookmark", "page", "B")).toBe("PAGEREF B \\h");
    expect(crossReferenceInstruction("bookmark", "aboveBelow", "B")).toBe("REF B \\p \\h");
    expect(crossReferenceInstruction("footnote", "number", "F")).toBe("NOTEREF F \\h");
    expect(crossReferenceInstruction("footnote", "page", "F")).toBe("PAGEREF F \\h");
    expect(crossReferenceInstruction("endnote", "number", "E")).toBe("NOTEREF E \\h");
    expect(crossReferenceInstruction("endnote", "aboveBelow", "E")).toBe("NOTEREF E \\p \\h");
    expect(crossReferenceInstruction("caption", "entire", "C")).toBe("REF C \\h");
    expect(crossReferenceInstruction("caption", "label", "C")).toBe("REF C \\h");
    expect(crossReferenceInstruction("caption", "captionText", "C")).toBe("REF C \\h");
    expect(crossReferenceInstruction("caption", "page", "C")).toBe("PAGEREF C \\h");
    // Equation defers to the equation workstream (G4).
    expect(crossReferenceInstruction("equation", "text", "X")).toBeNull();
    expect(CROSS_REFERENCE_CONTENTS.equation).toEqual([]);
  });

  it("disables content without data (heading number, caption text)", () => {
    const target = (extra: Partial<CrossReferenceTarget>): CrossReferenceTarget => ({
      key: "k",
      kind: "heading",
      name: "",
      pos: 0,
      text: "第一章",
      listText: "第一章",
      contentFrom: 1,
      contentTo: 4,
      ...extra,
    });
    const heading = target({});
    expect(crossReferenceContentAvailable(heading, "number")).toBe(false);
    expect(crossReferenceContentAvailable(heading, "text")).toBe(true);
    expect(crossReferenceContentAvailable(target({ number: "1" }), "number")).toBe(true);
    const caption = target({ kind: "caption", text: "Figure 1", captionFull: "Figure 1" });
    expect(crossReferenceContentAvailable(caption, "captionText")).toBe(false);
    expect(crossReferenceContentAvailable(caption, "entire")).toBe(true);
    expect(crossReferenceContentAvailable(target({ captionText: "Alpha" }), "captionText")).toBe(
      true,
    );
  });

  it("resolves REF \\n paragraph numbers and \\p above/below", () => {
    const bookmarks = new Map<string, FieldBookmark>([
      ["B", { text: "图 1", number: "2.1", page: 5, pageFormat: "lowerRoman", pos: 10 }],
    ]);
    const ctx = { bookmarks, selfPos: 20, positionTerms: { above: "上方", below: "下方" } };
    expect(resolveField("REF B \\n \\h", ctx)).toBe("2.1");
    // The target before the field reads "above"; after it, "below".
    expect(resolveField("REF B \\p \\h", ctx)).toBe("上方");
    expect(resolveField("REF B \\p \\h", { ...ctx, selfPos: 5 })).toBe("下方");
    expect(resolveField("PAGEREF B \\h", ctx)).toBe("v");
    expect(resolveField("PAGEREF B \\p \\h", ctx)).toBe("上方");
    expect(resolveField("NOTEREF B \\h", ctx)).toBe("图 1");
    expect(resolveField("NOTEREF B \\p \\h", ctx)).toBe("上方");
    // No bookmark / no field position / no `\n` number: keep the cache.
    expect(resolveField("REF 不存在", ctx)).toBeNull();
    expect(resolveField("REF B \\p", { bookmarks })).toBeNull();
    expect(
      resolveField("REF B \\n", { bookmarks: new Map([["B", { text: "x" }]]), selfPos: 1 }),
    ).toBeNull();
  });
});

describe("cross-reference commit matrix", () => {
  it("commits every heading content against a numbered heading", () => {
    const editor = build(
      [heading(1, "第一章", { numbering: { reference: "list_1", level: 0 } }), paragraph("refs")],
      { numbering: NUMBERING },
    );
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    for (const content of CROSS_REFERENCE_CONTENTS.heading) {
      const target = targetOfKind(dialogs, "heading");
      commit(dialogs, target.key, content);
      const field = fieldsOf(editor).at(-1)!;
      expect(field.instruction).toMatch(
        content === "number"
          ? /^REF _Ref\d{8} \\n \\h$/
          : content === "page"
            ? /^PAGEREF _Ref\d{8} \\h$/
            : content === "aboveBelow"
              ? /^REF _Ref\d{8} \\p \\h$/
              : /^REF _Ref\d{8} \\h$/,
      );
      expect(field.cached).not.toBe("");
    }
    // above/below: every target sits before the caret.
    expect(fieldsOf(editor).at(-1)!.cached).toBe("above");
    // The heading number is the live list number.
    expect(fieldsOf(editor)[2]!.cached).toBe("1");
    editor.destroy();
  });

  it("commits every numbered-item content", () => {
    const editor = build(
      [
        {
          type: "paragraph",
          attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
          content: [{ type: "text", text: "第一条" }],
        },
        paragraph("refs"),
      ],
      { numbering: NUMBERING },
    );
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    for (const content of CROSS_REFERENCE_CONTENTS.numbered) {
      const target = targetOfKind(dialogs, "numbered");
      commit(dialogs, target.key, content);
    }
    const fields = fieldsOf(editor);
    expect(fields.map((entry) => entry.instruction)).toEqual([
      "REF _Ref00000000 \\n \\h",
      "PAGEREF _Ref00000000 \\h",
      "REF _Ref00000000 \\p \\h",
    ]);
    // The generated reference's number comes from buildListLevels (compile
    // would register the same definition).
    expect(fields.map((entry) => entry.cached)).toEqual(["1", "1", "above"]);
    editor.destroy();
  });

  it("commits every bookmark content", () => {
    const editor = build([
      {
        type: "paragraph",
        content: [
          fieldAtom({ bookmarkStart: { id: 1, name: "目标" } }),
          { type: "text", text: "第三章" },
          fieldAtom({ bookmarkEnd: { id: 1 } }),
        ],
      },
      paragraph("refs"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    for (const content of CROSS_REFERENCE_CONTENTS.bookmark) {
      const target = targetOfKind(dialogs, "bookmark");
      commit(dialogs, target.key, content);
    }
    const fields = fieldsOf(editor);
    expect(fields.map((entry) => entry.instruction)).toEqual([
      "REF 目标 \\h",
      "PAGEREF 目标 \\h",
      "REF 目标 \\p \\h",
    ]);
    expect(fields.map((entry) => entry.cached)).toEqual(["第三章", "1", "above"]);
    editor.destroy();
  });

  it("commits every footnote and endnote content", () => {
    const editor = build([paragraph("a"), paragraph("b"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "a");
    dialogs.onNoteOk({ detail: { kind: "footnote", text: "脚注一" } } as unknown as Event);
    caretIn(editor, "b");
    dialogs.onNoteOk({ detail: { kind: "endnote", text: "尾注一" } } as unknown as Event);
    caretIn(editor, "refs");
    for (const content of CROSS_REFERENCE_CONTENTS.footnote) {
      const target = targetOfKind(dialogs, "footnote");
      commit(dialogs, target.key, content);
    }
    for (const content of CROSS_REFERENCE_CONTENTS.endnote) {
      const target = targetOfKind(dialogs, "endnote");
      commit(dialogs, target.key, content);
    }
    const fields = fieldsOf(editor);
    expect(fields.map((entry) => entry.instruction)).toEqual([
      "NOTEREF _Ref00000000 \\h",
      "PAGEREF _Ref00000000 \\h",
      "NOTEREF _Ref00000000 \\p \\h",
      "NOTEREF _Ref00000001 \\h",
      "PAGEREF _Ref00000001 \\h",
      "NOTEREF _Ref00000001 \\p \\h",
    ]);
    // Footnote 1, endnote i (Word's defaults), then above/below words.
    expect(fields.map((entry) => entry.cached)).toEqual(["1", "1", "above", "i", "1", "above"]);
    // The hidden bookmark wraps the reference atom itself.
    expect(bookmarksOf(editor).map((entry) => entry.name)).toContain("_Ref00000000");
    editor.destroy();
  });

  it("commits every caption content (entire, label, caption text, page, above/below)", () => {
    const editor = build([heading(1, "第一章"), paragraph("内容"), paragraph("refs")]);
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
        format: "ARABIC",
      },
    } as unknown as Event);
    caretIn(editor, "refs");
    for (const content of CROSS_REFERENCE_CONTENTS.caption) {
      const target = targetOfKind(dialogs, "caption");
      commit(dialogs, target.key, content);
    }
    const fields = fieldsOf(editor);
    // The caption's own SEQ field is first, the five references follow.
    const refs = fields.slice(1);
    expect(refs.map((entry) => entry.instruction)).toEqual([
      "REF _Ref00000001 \\h", // entire caption (its own hidden bookmark)
      "REF _Ref00000000 \\h", // label + number (reuses the caption's own pair)
      "REF _Ref00000002 \\h", // caption text (its own hidden pair)
      "PAGEREF _Ref00000000 \\h",
      "REF _Ref00000000 \\p \\h",
    ]);
    expect(refs.map((entry) => entry.cached)).toEqual([
      "Figure 1:1: Alpha chart",
      "Figure 1:1",
      "Alpha chart",
      "1",
      "above",
    ]);
    // The chapter-prefixed label+number is what "label and number" reads.
    expect(targetOfKind(dialogs, "caption").text).toBe("Figure 1:1");
    expect(targetOfKind(dialogs, "caption").captionText).toBe("Alpha chart");
    editor.destroy();
  });
});

describe("hidden bookmarks for headings", () => {
  it("creates one _Ref pair around the heading and reuses it", () => {
    const editor = build([heading(1, "第一章 概述"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    const key = targetOfKind(dialogs, "heading").key;
    // No bookmark yet.
    expect(bookmarksOf(editor)).toEqual([]);
    commit(dialogs, key, "text");
    let pairs = bookmarksOf(editor);
    expect(pairs).toHaveLength(1);
    expect(pairs[0]!.name).toBe("_Ref00000000");
    // The pair wraps the heading's text (paragraph at 0, start atom at 1).
    expect(pairs[0]!.from).toBe(2);
    expect(pairs[0]!.to).toBe(2 + "第一章 概述".length);
    // A second reference reuses the same pair.
    commit(dialogs, targetOfKind(dialogs, "heading").key, "aboveBelow");
    expect(bookmarksOf(editor)).toHaveLength(1);
    expect(fieldsOf(editor).map((entry) => entry.instruction)).toEqual([
      "REF _Ref00000000 \\h",
      "REF _Ref00000000 \\p \\h",
    ]);
    editor.destroy();
  });

  it("round-trips through JSON: the scan and update still resolve", () => {
    const editor = build([heading(1, "第一章"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    commit(dialogs, targetOfKind(dialogs, "heading").key, "text");
    const json = editor.getJSON();
    editor.destroy();

    const roundTripped = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: json,
    });
    const dialogs2 = new DialogCommands(host(roundTripped));
    const target = targetOfKind(dialogs2, "heading");
    expect(target.name).toBe("_Ref00000000");
    expect(target.text).toBe("第一章");
    expect(dialogs2.updateAllFields()).toBe(0);
    roundTripped.destroy();
  });

  it("refreshes REF text when the heading changes", () => {
    const editor = build([heading(1, "旧标题"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    commit(dialogs, targetOfKind(dialogs, "heading").key, "text");
    expect(fieldsOf(editor)[0]!.cached).toBe("旧标题");

    editor.commands.command(({ state, dispatch }) => {
      let done = false;
      state.doc.descendants((node, pos) => {
        if (done || node.type.name !== "paragraph" || !node.attrs.heading) return true;
        dispatch?.(state.tr.insertText("新标题", pos + 2, pos + 2 + node.textContent.length));
        done = true;
        return false;
      });
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor)[0]!.cached).toBe("新标题");
    editor.destroy();
  });
});

describe("PAGEREF and above/below updates", () => {
  it("follows the target's page when content before it moves", () => {
    const editor = build([
      {
        type: "paragraph",
        content: [
          fieldAtom({ bookmarkStart: { id: 1, name: "目标" } }),
          { type: "text", text: "第三章" },
          fieldAtom({ bookmarkEnd: { id: 1 } }),
        ],
      },
      paragraph("refs"),
    ]);
    const pages = new Map([
      ["第三章", 1],
      ["refs", 1],
    ]);
    const dialogs = new DialogCommands(host(editor, { frame: frameByText(editor, pages) }));
    caretIn(editor, "refs");
    commit(dialogs, targetOfKind(dialogs, "bookmark").key, "page");
    expect(fieldsOf(editor)[0]!.instruction).toBe("PAGEREF 目标 \\h");
    expect(fieldsOf(editor)[0]!.cached).toBe("1");

    // A page break landed before the bookmark: the target is on page 3 now.
    pages.set("第三章", 3);
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor)[0]!.cached).toBe("3");
    editor.destroy();
  });

  it("flips above/below when the reference moves before the target", () => {
    // The bookmark must be a real pair (a bare paragraph has no user target).
    const editor = build([
      {
        type: "paragraph",
        content: [
          fieldAtom({ bookmarkStart: { id: 1, name: "目标" } }),
          { type: "text", text: "第三章" },
          fieldAtom({ bookmarkEnd: { id: 1 } }),
        ],
      },
      paragraph("refs"),
    ]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    commit(dialogs, targetOfKind(dialogs, "bookmark").key, "aboveBelow");
    expect(fieldsOf(editor)[0]!.cached).toBe("above");

    // Move the reference paragraph before the bookmarked one.
    editor.commands.command(({ state, dispatch }) => {
      const last = state.doc.lastChild!;
      const tr = state.tr.delete(state.doc.content.size - last.nodeSize, state.doc.content.size);
      tr.insert(0, last);
      dispatch?.(tr);
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor)[0]!.cached).toBe("below");
    editor.destroy();
  });
});

describe("numbered-item references", () => {
  it("renders the live list number and renumbers when a list item is inserted before", () => {
    const editor = build(
      [
        {
          type: "paragraph",
          attrs: { numbering: { reference: "list_1", level: 0 } },
          content: [{ type: "text", text: "第一条" }],
        },
        {
          type: "paragraph",
          attrs: { numbering: { reference: "list_1", level: 0 } },
          content: [{ type: "text", text: "第二条" }],
        },
        paragraph("refs"),
      ],
      { numbering: NUMBERING },
    );
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "numbered", "第二条");
    expect(target.number).toBe("2");
    commit(dialogs, target.key, "number");
    expect(fieldsOf(editor)[0]!.instruction).toBe("REF _Ref00000000 \\n \\h");
    expect(fieldsOf(editor)[0]!.cached).toBe("2");

    // Insert a new first list item: the referenced paragraph becomes 3.
    editor.commands.command(({ state, dispatch }) => {
      const node = state.schema.nodeFromJSON({
        type: "paragraph",
        attrs: { numbering: { reference: "list_1", level: 0 } },
        content: [{ type: "text", text: "新的第一条" }],
      });
      dispatch?.(state.tr.insert(0, node));
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor)[0]!.cached).toBe("3");
    editor.destroy();
  });

  it("builds multi-level numbers from the level text", () => {
    const editor = build(
      [
        {
          type: "paragraph",
          attrs: { numbering: { reference: "list_1", level: 0 } },
          content: [{ type: "text", text: "第一条" }],
        },
        {
          type: "paragraph",
          attrs: { numbering: { reference: "list_1", level: 1 } },
          content: [{ type: "text", text: "子条" }],
        },
        paragraph("refs"),
      ],
      { numbering: NUMBERING },
    );
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "numbered", "子条");
    expect(target.number).toBe("1.a");
    commit(dialogs, target.key, "number");
    expect(fieldsOf(editor)[0]!.cached).toBe("1.a");
    expect(dialogs.updateAllFields()).toBe(0);
    editor.destroy();
  });

  it("uses a style-numbered heading's live number when it carries direct numbering", () => {
    const editor = build(
      [
        heading(1, "第一章"),
        heading(2, "第一节", { numbering: { reference: "list_1", level: 0 } }),
        paragraph("refs"),
      ],
      { numbering: NUMBERING },
    );
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "heading", "第一节");
    expect(target.number).toBe("1");
    commit(dialogs, target.key, "number");
    expect(fieldsOf(editor)[0]!.instruction).toBe("REF _Ref00000000 \\n \\h");
    expect(fieldsOf(editor)[0]!.cached).toBe("1");
    editor.destroy();
  });
});

describe("note reference numbers", () => {
  it("renumbers NOTEREF after the first note is deleted", () => {
    const editor = build([paragraph("a"), paragraph("b"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "a");
    dialogs.onNoteOk({ detail: { kind: "footnote", text: "第一" } } as unknown as Event);
    caretIn(editor, "b");
    dialogs.onNoteOk({ detail: { kind: "footnote", text: "第二" } } as unknown as Event);
    caretIn(editor, "refs");

    const second = dialogs
      .crossReferenceTargets()
      .filter((entry) => entry.kind === "footnote")
      .at(-1)!;
    expect(second.number).toBe("2");
    commit(dialogs, second.key, "number");
    expect(fieldsOf(editor)[0]!.instruction).toBe("NOTEREF _Ref00000000 \\h");
    expect(fieldsOf(editor)[0]!.cached).toBe("2");

    // Delete the first note's reference atom: the second becomes note 1.
    editor.commands.command(({ state, dispatch }) => {
      let at: number | null = null;
      state.doc.descendants((node, pos) => {
        if (at != null || node.type.name !== "inlinePassthrough") return true;
        try {
          const data = JSON.parse(String(node.attrs.data ?? "{}")) as {
            footnoteReference?: number;
          };
          if (data.footnoteReference === 1) at = pos;
        } catch {
          /* opaque */
        }
        return true;
      });
      if (at == null) return false;
      dispatch?.(state.tr.delete(at, at + 1));
      return true;
    });
    expect(dialogs.updateAllFields()).toBe(1);
    expect(fieldsOf(editor)[0]!.cached).toBe("1");
    editor.destroy();
  });

  it("keeps the endnote's Roman numeral shape", () => {
    const editor = build([paragraph("a"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "a");
    dialogs.onNoteOk({ detail: { kind: "endnote", text: "尾注一" } } as unknown as Event);
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "endnote");
    expect(target.number).toBe("i");
    commit(dialogs, target.key, "number");
    expect(fieldsOf(editor)[0]!.cached).toBe("i");
    expect(dialogs.updateAllFields()).toBe(0);
    editor.destroy();
  });
});

describe("dialog-option guards", () => {
  it("does not insert a heading number reference for an unnumbered heading", () => {
    const editor = build([heading(1, "第一章"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "heading");
    expect(target.number).toBeUndefined();
    commit(dialogs, target.key, "number");
    expect(fieldsOf(editor)).toEqual([]);
    editor.destroy();
  });

  it("does not insert an entire-caption reference for a caption without text", () => {
    const editor = build([paragraph("内容"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "", position: "below" },
    } as unknown as Event);
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "caption");
    expect(target.captionText).toBeUndefined();
    const before = fieldsOf(editor).length;
    commit(dialogs, target.key, "captionText");
    expect(fieldsOf(editor)).toHaveLength(before);
    editor.destroy();
  });
});

describe("caption text separator handling (D1)", () => {
  it("keeps 'Only caption text' separator-free across updates for every separator shape", () => {
    for (const raw of [
      ": Alpha chart",
      "- Alpha chart",
      ". Alpha chart",
      "\u2014 Alpha chart",
      "：Alpha chart",
    ]) {
      const editor = build([
        {
          type: "paragraph",
          attrs: { style: "Caption" },
          content: [
            fieldAtom({ bookmarkStart: { id: 5, name: "_Ref00000005" } }),
            { type: "text", text: "Figure " },
            fieldAtom({
              simpleField: { instruction: "SEQ Figure \\* ARABIC", cachedValue: "1" },
            }),
            fieldAtom({ bookmarkEnd: { id: 5 } }),
            { type: "text", text: raw },
          ],
        },
        paragraph("refs"),
      ]);
      const dialogs = new DialogCommands(host(editor));
      caretIn(editor, "refs");
      const target = targetOfKind(dialogs, "caption");
      expect(target.captionText).toBe("Alpha chart");
      commit(dialogs, target.key, "captionText");
      expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
      // One pass and a second pass must both keep the separator out of the
      // re-derived bookmark text.
      expect(dialogs.updateAllFields()).toBe(0);
      expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
      expect(dialogs.updateAllFields()).toBe(0);
      expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
      editor.destroy();
    }
  });

  it("keeps the A1 caption dialog's colon out across F9 and Update All", () => {
    const editor = build([paragraph("内容"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Alpha chart", position: "below" },
    } as unknown as Event);
    caretIn(editor, "refs");
    const target = targetOfKind(dialogs, "caption");
    commit(dialogs, target.key, "captionText");
    expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
    // F9 on the reference itself.
    editor.commands.setTextSelection(fieldsOf(editor).at(-1)!.pos + 1);
    dialogs.fieldUpdateAtSelection();
    expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
    // Update All Fields, twice.
    expect(dialogs.updateAllFields()).toBe(0);
    expect(dialogs.updateAllFields()).toBe(0);
    expect(fieldsOf(editor).at(-1)!.cached).toBe("Alpha chart");
    editor.destroy();
  });
});

describe("Update All refreshes references to a renumbered caption (D2)", () => {
  it("applies the fresh SEQ number to label/entire references in one pass", () => {
    const editor = build([heading(1, "第一章"), paragraph("内容"), paragraph("refs")]);
    const dialogs = new DialogCommands(host(editor));
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Alpha", position: "below" },
    } as unknown as Event);
    caretIn(editor, "refs");
    commit(dialogs, targetOfKind(dialogs, "caption").key, "label");
    commit(dialogs, targetOfKind(dialogs, "caption").key, "entire");
    const refValues = (): string[] =>
      fieldsOf(editor)
        .filter((entry) => entry.instruction.startsWith("REF _Ref"))
        .map((entry) => entry.cached);
    expect(refValues()).toEqual(["Figure 1", "Figure 1: Alpha"]);

    // A new caption above the referenced one pushes it to Figure 2; the
    // existing references still cache Figure 1.
    caretIn(editor, "内容");
    dialogs.onCaptionOk({
      detail: { label: "Figure", text: "Beta", position: "above" },
    } as unknown as Event);
    expect(refValues()).toEqual(["Figure 1", "Figure 1: Alpha"]);

    let transactions = 0;
    let docChanged = 0;
    editor.on("transaction", ({ transaction }) => {
      transactions += 1;
      if (transaction.docChanged) docChanged += 1;
    });

    // ONE pass: the caption's SEQ renumber lands before the REF-family
    // resolves, so the references quote the new number immediately.
    expect(dialogs.updateAllFields()).toBe(3);
    expect(docChanged).toBe(1);
    expect(transactions).toBe(1);
    expect(refValues()).toEqual(["Figure 2", "Figure 2: Alpha"]);
    expect(dialogs.updateAllFields()).toBe(0);
    expect(refValues()).toEqual(["Figure 2", "Figure 2: Alpha"]);
    editor.destroy();
  });
});
