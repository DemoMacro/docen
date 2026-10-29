// @vitest-environment happy-dom
import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { CommentsCommands } from "./comments";

const Text = TextNode.create({ name: "text", group: "inline" });

const marker = (id: number): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify({ commentRangeStart: { id } }) },
  }) as JSONContent;

const buildEditor = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      attrs: {
        documentExtras: {
          comments: [
            {
              id: 1,
              author: "Alice",
              children: [{ children: [{ text: "root" }], paraId: "AAAAAA01" }],
            },
            {
              id: 2,
              author: "Alice",
              children: [{ children: [{ text: "reply" }], paraId: "AAAAAA02" }],
            },
            { id: 3, author: "Bob", children: [{ children: [{ text: "other" }] }] },
          ],
          commentsExtended: [
            { paraId: "AAAAAA01", done: false },
            { paraId: "AAAAAA02", paraIdParent: "AAAAAA01", done: false },
          ],
        },
      },
      content: [
        {
          type: "paragraph",
          content: [marker(1), { type: "text", text: "reviewed" }, marker(3)],
        },
      ],
    },
  });

const host = (editor: EditorType): CommentsCommands =>
  new CommentsCommands({
    editor: () => editor,
    bridge: () => undefined,
    element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
    showTaskpane: () => {},
  });

const extrasOf = (
  editor: EditorType,
): {
  comments: Array<{ id?: number }>;
  commentsExtended: Array<{ paraId?: string }>;
} => editor.state.doc.attrs.documentExtras as never;

const markerIds = (editor: EditorType): number[] => {
  const ids: number[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "inlinePassthrough") return true;
    const data = JSON.parse(String(node.attrs.data)) as { commentRangeStart?: { id?: number } };
    if (typeof data.commentRangeStart?.id === "number") ids.push(data.commentRangeStart.id);
    return true;
  });
  return ids;
};

describe("CommentsCommands comment deletion", () => {
  it("removes a thread root with its reply and round-trip extras", () => {
    const editor = buildEditor();
    const commands = host(editor);

    commands.onCommentDelete(new CustomEvent("comment:delete", { detail: { id: 1 } }));

    expect(markerIds(editor)).toEqual([3]);
    expect(extrasOf(editor).comments.map(({ id }) => id)).toEqual([3]);
    expect(extrasOf(editor).commentsExtended).toEqual([]);
  });
});
