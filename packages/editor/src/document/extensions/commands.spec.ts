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

type AnyNode = {
  attrs: Record<string, unknown>;
  childCount: number;
  child: (i: number) => AnyNode;
};

/** The first node of `name` in document order. `descendants` cannot abort the
 *  whole walk (a falsy return only skips children), so collect and take [0]. */
const firstNodeOf = (editor: EditorType, name: string): AnyNode => {
  const found: AnyNode[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === name) {
      found.push(node as unknown as AnyNode);
      return false;
    }
    return true;
  });
  return found[0]!;
};

/** Move the caret into cell (row, col) of the first table — inside its first
 *  paragraph, the way a user's caret sits. */
const caretInCell = (editor: EditorType, row = 0, col = 0): void => {
  const table = tablesOf(editor)[0]!;
  let cellPos = -1;
  editor.state.doc.descendants((node, nodePos) => {
    if (node.type.name !== "table") return true;
    let rowPos = nodePos + 1;
    for (let r = 0; r < row; r += 1) rowPos += node.child(r).nodeSize;
    cellPos = rowPos + 1;
    const rowNode = node.child(row);
    for (let c = 0; c < col; c += 1) cellPos += rowNode.child(c).nodeSize;
    return false;
  });
  editor.commands.setTextSelection(cellPos + 2);
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

describe("table row / column commands", () => {
  it("insert-row-above/below add rows around the caret's row", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 2, 0);
    expect(editor.commands["insert-row-above"]()).toBe(true);
    expect(tablesOf(editor)[0]!.childCount).toBe(4);
    caretInCell(editor, 1, 0);
    expect(editor.commands["insert-row-below"]()).toBe(true);
    expect(tablesOf(editor)[0]!.childCount).toBe(5);
  });

  it("insert-column-right extends every row", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["insert-column-right"]()).toBe(true);
    const table = tablesOf(editor)[0]!;
    expect(table.childCount).toBe(3);
    for (let r = 0; r < 3; r += 1) expect(table.child(r).childCount).toBe(4);
  });

  it("delete-row removes the caret's row; the last row deletes the table", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 2, 0);
    expect(editor.commands["delete-row"]()).toBe(true);
    expect(tablesOf(editor)[0]!.childCount).toBe(2);
    caretInCell(editor, 0, 0);
    expect(editor.commands["delete-row"]()).toBe(true);
    expect(tablesOf(editor)[0]!.childCount).toBe(1);
    caretInCell(editor, 0, 0);
    expect(editor.commands["delete-row"]()).toBe(true);
    expect(tablesOf(editor)).toHaveLength(0);
  });

  it("delete-column removes a column; the last column deletes the table", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 2);
    expect(editor.commands["delete-column"]()).toBe(true);
    expect(tablesOf(editor)[0]!.child(0).childCount).toBe(2);
    caretInCell(editor, 0, 0);
    expect(editor.commands["delete-column"]()).toBe(true);
    expect(tablesOf(editor)[0]!.child(0).childCount).toBe(1);
    caretInCell(editor, 0, 0);
    expect(editor.commands["delete-column"]()).toBe(true);
    expect(tablesOf(editor)).toHaveLength(0);
  });
});

describe("table cell property commands", () => {
  it("align-cell stamps the cell's verticalAlign and its paragraphs' alignment", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["align-cell"]("bc")).toBe(true);
    const cell = firstNodeOf(editor, "tableCell");
    expect(cell.attrs.verticalAlign).toBe("bottom");
    expect(cell.child(0).attrs.alignment).toBe("center");
    expect(editor.commands["align-cell"]("bogus")).toBe(false);
  });

  it("repeat-header-rows toggles the row's tblHeader flag", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 1, 0);
    expect(tablesOf(editor)[0]!.child(1).attrs.tableHeader).toBeFalsy();
    expect(editor.commands["repeat-header-rows"]()).toBe(true);
    expect(tablesOf(editor)[0]!.child(1).attrs.tableHeader).toBe(true);
    expect(editor.commands["repeat-header-rows"]()).toBe(true);
    expect(tablesOf(editor)[0]!.child(1).attrs.tableHeader).toBe(false);
  });

  it("cell-shading stamps the cell shading, clearing with none", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["cell-shading"]("FFEE88")).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.shading).toEqual({
      fill: "FFEE88",
      type: "clear",
    });
    expect(editor.commands["cell-shading"]("none")).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.shading).toBeFalsy();
  });

  it("table-style applies a border preset to the enclosing table", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["table-style"]("light-list")).toBe(true);
    const borders = tablesOf(editor)[0]!.attrs.borders as Record<string, unknown>;
    expect(borders.top).toEqual(GRID);
    expect(borders.insideVertical).toEqual({ style: "none", size: 0, color: "auto" });
    expect(editor.commands["table-style"]("bogus")).toBe(false);
  });

  it("text-direction toggles the cell's tcPr textDirection", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(firstNodeOf(editor, "tableCell").attrs.textDirection).toBeFalsy();
    expect(editor.commands["text-direction"]()).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.textDirection).toBe("tbRl");
    expect(editor.commands["text-direction"]()).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.textDirection).toBeFalsy();
  });
});

describe("select-table-column / convert-to-text", () => {
  it("select-table-column selects the caret's column across all rows", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 1); // middle cell of the header row
    expect(editor.commands["select-table-column"]()).toBe(true);
    const { from, to, empty } = editor.state.selection;
    expect(empty).toBe(false);
    // The selection spans from the header row's cell into the last row's.
    const $from = editor.state.doc.resolve(from);
    const $to = editor.state.doc.resolve(to);
    expect($from.node(2).type.name).toBe("tableRow");
    expect($to.node(2).type.name).toBe("tableRow");
    expect($from.before(2)).toBeLessThan($to.before(2));
  });

  it("convert-to-text replaces the table with tab-joined paragraphs", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    // Type into the first cell so the conversion has content to move.
    editor.commands.setContent({
      type: "doc",
      content: [
        {
          type: "table",
          content: [
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "甲" }] }],
                },
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "乙" }] }],
                },
              ],
            },
          ],
        },
      ],
    });
    caretInCell(editor, 0, 0);
    expect(editor.commands["convert-to-text"]()).toBe(true);
    expect(tablesOf(editor)).toHaveLength(0);
    expect(editor.state.doc.childCount).toBe(1);
    expect(editor.state.doc.firstChild?.type.name).toBe("paragraph");
    expect(editor.state.doc.textContent).toBe("甲\t乙");
  });
});
