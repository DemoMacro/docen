// @vitest-environment node
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { CellSelection, cellAt, inSameTable } from "./cell-selection";

// Tiptap's schema needs the plain text node (same trick as the other specs).
const Text = TextNode.create({ name: "text", group: "inline" });

const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Table, TableRow, TableCell],
    content: {
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
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "丙" }] }],
                },
              ],
            },
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "丁" }] }],
                },
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "戊" }] }],
                },
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "己" }] }],
                },
              ],
            },
          ],
        },
        { type: "paragraph" },
      ],
    },
  });

/** The cell (row, col) start position of the first table. */
const cellPos = (editor: EditorType, row: number, col: number): number => {
  let pos = -1;
  editor.state.doc.descendants((node, nodePos) => {
    if (node.type.name !== "table") return true;
    let rowPos = nodePos + 1;
    for (let r = 0; r < row; r += 1) rowPos += node.child(r).nodeSize;
    const rowNode = node.child(row);
    pos = rowPos + 1;
    for (let c = 0; c < col; c += 1) pos += rowNode.child(c).nodeSize;
    return false;
  });
  return pos;
};

const cellsOf = (editor: EditorType, selection: CellSelection): string[] => {
  const texts: string[] = [];
  selection.forEachCell((node) => texts.push(node.textContent));
  return texts;
};

describe("cellAt / inSameTable", () => {
  it("resolves a position inside a cell to that cell; null outside", () => {
    const editor = build();
    const { doc } = editor.state;
    const inCell = doc.resolve(cellPos(editor, 0, 1) + 2); // inside 乙's paragraph
    const cell = cellAt(inCell);
    expect(cell?.pos).toBe(cellPos(editor, 0, 1));
    expect(cellAt(doc.resolve(0))).toBeNull();
    // Both cells of one table share it; a body paragraph does not.
    const a = cellAt(doc.resolve(cellPos(editor, 0, 0) + 1))!;
    const b = cellAt(doc.resolve(cellPos(editor, 1, 2) + 1))!;
    expect(inSameTable(a, b)).toBe(true);
    expect(inSameTable(a, inCell.doc.resolve(editor.state.doc.content.size - 1))).toBe(false);
  });
});

describe("CellSelection rectangles", () => {
  it("drag rect selects every cell whole, in document order", () => {
    const editor = build();
    const { doc } = editor.state;
    const sel = CellSelection.create(doc, cellPos(editor, 0, 0), cellPos(editor, 1, 2));
    expect(cellsOf(editor, sel)).toEqual(["甲", "乙", "丙", "丁", "戊", "己"]);
  });

  it("a reversed drag selects the same rectangle", () => {
    const editor = build();
    const { doc } = editor.state;
    const sel = CellSelection.create(doc, cellPos(editor, 1, 2), cellPos(editor, 0, 0));
    expect(cellsOf(editor, sel)).toEqual(["甲", "乙", "丙", "丁", "戊", "己"]);
  });

  it("a single-cell drag selects that cell only", () => {
    const editor = build();
    const { doc } = editor.state;
    const sel = CellSelection.create(doc, cellPos(editor, 1, 1));
    expect(cellsOf(editor, sel)).toEqual(["戊"]);
  });

  it("rowSelection takes the cell's row; colSelection its column", () => {
    const editor = build();
    const { doc } = editor.state;
    const row = CellSelection.rowSelection(doc.resolve(cellPos(editor, 0, 1)));
    expect(cellsOf(editor, row)).toEqual(["甲", "乙", "丙"]);
    const col = CellSelection.colSelection(doc.resolve(cellPos(editor, 1, 2)));
    expect(cellsOf(editor, col)).toEqual(["丙", "己"]);
  });

  it("survives a transaction map and clears cells on replace", () => {
    const editor = build();
    const { doc } = editor.state;
    const sel = CellSelection.create(doc, cellPos(editor, 0, 0), cellPos(editor, 1, 1));
    // Delete the paragraph before the table: the selection must keep its
    // cell shape (the grid shift maps both ends back into cells).
    editor.commands.command(({ tr, dispatch }) => {
      tr.delete(0, 1);
      tr.setSelection(sel.map(tr.doc, tr.mapping) as never);
      dispatch?.(tr);
      return true;
    });
    const mapped = editor.state.selection as unknown as CellSelection;
    expect(mapped).toBeInstanceOf(CellSelection);
    expect(cellsOf(editor, mapped)).toEqual(["甲", "乙", "丁", "戊"]);
    // Word's cell delete: content empties, the grid survives.
    editor.commands.command(({ tr, dispatch }) => {
      mapped.replace(tr as never);
      dispatch?.(tr);
      return true;
    });
    const tableTexts: string[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell") tableTexts.push(node.textContent);
      return true;
    });
    expect(tableTexts).toEqual(["", "", "丙", "", "", "己"]);
  });

  it("preserves cell attributes such as shading and columnSpan on replace", () => {
    const editor = build();
    // Set custom shading on cell (0, 0).
    const cPos = cellPos(editor, 0, 0);
    editor.commands.command(({ tr, dispatch }) => {
      const node = tr.doc.nodeAt(cPos)!;
      tr.setNodeMarkup(cPos, undefined, { ...node.attrs, shading: "FFFF00", columnSpan: 2 });
      dispatch?.(tr);
      return true;
    });
    const sel = CellSelection.create(editor.state.doc, cPos, cPos);
    editor.commands.command(({ tr, dispatch }) => {
      sel.replace(tr as never);
      dispatch?.(tr);
      return true;
    });
    const clearedCell = editor.state.doc.nodeAt(cPos)!;
    expect(clearedCell.textContent).toBe("");
    expect(clearedCell.attrs.shading).toBe("FFFF00");
    expect(clearedCell.attrs.columnSpan).toBe(2);
  });
});
