// @vitest-environment node
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DocumentCommands } from "./commands";

// Tiptap's schema needs the plain text node (same trick as the TOC spec).
const Text = TextNode.create({ name: "text", group: "inline" });

/** A headless editor with exactly the schema the table commands touch. */
const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Table, TableRow, TableCell, DocumentCommands],
    content: { type: "doc", content: [{ type: "paragraph" }] },
  });

const tablesOf = (
  editor: EditorType,
): Array<{
  attrs: Record<string, unknown>;
  childCount: number;
  child: (i: number) => { attrs: Record<string, unknown>; childCount: number };
}> => {
  const out: Array<never> = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "table") out.push(node as never);
  });
  return out as never;
};

const GRID = { style: "single", size: 4, color: "auto" };
const GRID_BORDERS = {
  top: GRID,
  bottom: GRID,
  left: GRID,
  right: GRID,
  insideHorizontal: GRID,
  insideVertical: GRID,
};

describe("insert-table", () => {
  it("stamps Word's Table Grid borders and a header row, caret in the first cell", () => {
    const editor = build();
    expect(editor.commands["insert-table"]()).toBe(true);

    const tables = tablesOf(editor);
    expect(tables).toHaveLength(1);
    expect(tables[0]!.attrs.borders).toEqual(GRID_BORDERS);
    expect(tables[0]!.childCount).toBe(3);
    expect(tables[0]!.child(0).attrs.tableHeader).toBe(true);
    expect(tables[0]!.child(0).childCount).toBe(3);
    // The caret lands inside the first cell, ready to type (Word behavior).
    expect(editor.state.selection.from).toBeGreaterThan(0);
  });
});

describe("delete-table", () => {
  it("removes the enclosing table", () => {
    const editor = build();
    editor.commands["insert-table"]();
    expect(tablesOf(editor)).toHaveLength(1);

    expect(editor.commands["delete-table"]()).toBe(true);
    expect(tablesOf(editor)).toHaveLength(0);
  });

  it("is a no-op outside a table", () => {
    const editor = build();
    expect(editor.commands["delete-table"]()).toBe(false);
    expect(editor.state.doc.firstChild?.type.name).toBe("paragraph");
  });

  it("in a nested table removes only the inner one", () => {
    const editor = build();
    editor.commands["insert-table"]();
    // The caret sits in the first cell — the second insert nests inside it.
    editor.commands["insert-table"]();
    expect(tablesOf(editor)).toHaveLength(2);

    expect(editor.commands["delete-table"]()).toBe(true);
    const tables = tablesOf(editor);
    expect(tables).toHaveLength(1);
    // The survivor is the outer table.
    expect(tables[0]!.attrs.borders).toEqual(GRID_BORDERS);
  });
});
