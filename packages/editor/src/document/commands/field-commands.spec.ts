import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import type { FieldFrame } from "../fields";
import { DialogCommands, type DialogsHost } from "./dialogs";

const fieldAtom = (branch: object): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify(branch) },
  }) as JSONContent;

const Text = TextNode.create({ name: "text", group: "inline" });

const field = (instruction: string, cachedValue: string): JSONContent =>
  fieldAtom({ simpleField: { instruction, cachedValue } });

/** The doc's paragraph/atom positions for bookmark lookups — a constant page
 *  map keeps the spec independent of ProseMirror position arithmetic. */
const buildDoc = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      attrs: {
        core: { creator: "作者甲", title: "年度报告", revision: 7 },
        documentExtras: { customProperties: [{ name: "项目代码", value: "PRJ-1" }] },
      },
      content: [
        {
          type: "paragraph",
          content: [
            field("AUTHOR", "旧作者"),
            field("DOCPROPERTY 项目代码", ""),
            field("FILENAME", "旧.docx"),
            field("TITLE", ""),
          ],
        },
        { type: "paragraph", content: [field("SEQ 图 \\* ARABIC", "9")] },
        { type: "paragraph", content: [field("SEQ 图 \\* ARABIC", "9")] },
        {
          type: "paragraph",
          content: [
            fieldAtom({ bookmarkStart: { id: 1, name: "目标" } }),
            { type: "text", text: "第三章" },
            fieldAtom({ bookmarkEnd: { id: 1 } }),
            field("REF 目标 \\h", ""),
            field("PAGEREF 目标 \\h", ""),
          ],
        },
      ],
    },
  });

const host = (
  editor: EditorType,
  opts: { frame?: (pos: number) => FieldFrame | undefined; pageOf?: () => number | null } = {},
): DialogsHost => ({
  editor: () => editor,
  bridge: () => ({
    activeEditor: () => editor,
    focus: () => {},
    pageOf: () => opts.pageOf?.() ?? 2,
  }),
  element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
  syncStatusLanguage: () => {},
  filename: () => "报告.docx",
  fieldFrame: opts.frame,
});

const dataOf = (editor: EditorType, index: number): Record<string, unknown> => {
  const hits: string[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "inlinePassthrough") hits.push(String(node.attrs.data));
    return true;
  });
  return JSON.parse(hits[index]!) as Record<string, unknown>;
};

const simpleValue = (editor: EditorType, index: number): string | undefined =>
  (dataOf(editor, index).simpleField as { cachedValue?: string } | undefined)?.cachedValue;

describe("updateAllFields", () => {
  it("re-derives every field in one transaction", () => {
    const editor = buildDoc();
    const dialogs = new DialogCommands(host(editor));
    let transactions = 0;
    let docChanged = 0;
    editor.on("transaction", ({ transaction }) => {
      transactions += 1;
      if (transaction.docChanged) docChanged += 1;
    });
    const updated = dialogs.updateAllFields();
    expect(updated).toBe(8);
    expect(docChanged).toBe(1);
    expect(transactions).toBe(1);
    // Document information now reads the core attrs (the stale context bug).
    expect(simpleValue(editor, 0)).toBe("作者甲");
    expect(simpleValue(editor, 1)).toBe("PRJ-1");
    expect(simpleValue(editor, 2)).toBe("报告.docx");
    expect(simpleValue(editor, 3)).toBe("年度报告");
    // SEQ ordinals restart per label in document order.
    expect(simpleValue(editor, 4)).toBe("1");
    expect(simpleValue(editor, 5)).toBe("2");
    // REF reads the bookmark text; PAGEREF the bookmark's pinned page.
    expect(simpleValue(editor, 8)).toBe("第三章");
    expect(simpleValue(editor, 9)).toBe("3");
  });

  it("leaves page-dependent fields cached without a page frame", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [{ type: "paragraph", content: [field("PAGE", "7")] }],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(0);
    expect(simpleValue(editor, 0)).toBe("7");
  });

  it("resolves page-dependent fields through the host's page frame", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [field("PAGE", "7"), field("NUMPAGES", "7"), field("SECTION", "1")],
          },
        ],
      },
    });
    const frame = (): FieldFrame => ({
      page: 3,
      pageCount: 10,
      section: 2,
      sectionPages: 4,
      pageFormat: "lowerRoman",
    });
    const dialogs = new DialogCommands(host(editor, { frame }));
    expect(dialogs.updateAllFields()).toBe(3);
    expect(simpleValue(editor, 0)).toBe("iii");
    expect(simpleValue(editor, 1)).toBe("10");
    expect(simpleValue(editor, 2)).toBe("ii");
  });

  it("resolves PAGEREF with the bookmark's own section restart and numFmt", () => {
    // The bookmark paragraph sits in a section showing page 5 in lowercase
    // Roman; the PAGEREF field itself sits earlier (page 1, decimal). The
    // update must write the TARGET's "v", not the physical/own-section "1".
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [field("PAGEREF 目标 \\h", "1")] },
          {
            type: "paragraph",
            content: [
              fieldAtom({ bookmarkStart: { id: 1, name: "目标" } }),
              { type: "text", text: "第三章" },
              fieldAtom({ bookmarkEnd: { id: 1 } }),
            ],
          },
        ],
      },
    });
    let bookmarkPos = 0;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "paragraph" && node.textContent === "第三章") bookmarkPos = pos;
      return true;
    });
    const frame = (pos: number): FieldFrame | undefined =>
      pos === bookmarkPos
        ? { page: 5, pageCount: 9, section: 2, sectionPages: 4, pageFormat: "lowerRoman" }
        : { page: 1, pageCount: 9, section: 1, sectionPages: 5 };
    const dialogs = new DialogCommands(host(editor, { frame }));
    expect(dialogs.updateAllFields()).toBe(1);
    expect(simpleValue(editor, 0)).toBe("v");
  });

  it("keeps a malformed revision cached instead of painting NaN", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        attrs: { core: { revision: "not-a-number" } },
        content: [{ type: "paragraph", content: [field("REVNUM", "7")] }],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(0);
    expect(simpleValue(editor, 0)).toBe("7");
  });

  it("is idempotent — the on-open hook cannot loop renders", () => {
    const editor = buildDoc();
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.updateAllFields()).toBe(8);
    expect(dialogs.updateAllFields()).toBe(0);
  });
});

describe("fieldUpdateAtSelection", () => {
  it("re-derives the field under the caret from the page frame", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "ab" }, field("PAGE", "7")] },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor, { frame: () => ({ page: 4, pageCount: 9 }) }));
    editor.commands.setTextSelection(4); // right after the atom
    dialogs.fieldUpdateAtSelection();
    expect(simpleValue(editor, 0)).toBe("4");
  });

  it("keeps the cache when the context cannot evaluate the field", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "ab" }, field("PAGE", "7")] },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.fieldUpdateAtSelection();
    expect(simpleValue(editor, 0)).toBe("7");
  });
});
