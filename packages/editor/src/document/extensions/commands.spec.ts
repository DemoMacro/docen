// @vitest-environment node
import {
  Document,
  Image,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  WpgGroup,
  WpsShape,
} from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { NodeSelection } from "@tiptap/pm/state";
import { describe, expect, it } from "vitest";

import { CellSelection } from "../canvas/cell-selection";
import { DocumentCommands, listLevelStepPatch, normalizeParagraphAlignment } from "./commands";

// Tiptap's schema needs the plain text node (same trick as the TOC spec).
const Text = TextNode.create({ name: "text", group: "inline" });

/** The schema the document commands touch. */
const EXTENSIONS = [
  Document,
  Paragraph,
  Text,
  Table,
  TableRow,
  TableCell,
  Image,
  WpsShape,
  WpgGroup,
  DocumentCommands,
];

/** A headless editor with exactly the schema the table commands touch. */
const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: { type: "doc", content: [{ type: "paragraph" }] },
  });

type AnyNode = {
  attrs: Record<string, unknown>;
  childCount: number;
  child: (i: number) => AnyNode;
  textContent: string;
};

const tablesOf = (editor: EditorType): AnyNode[] => {
  const out: AnyNode[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "table") out.push(node as unknown as AnyNode);
  });
  return out;
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

/** A TextSelection dragging across cells (rowA,colA) → (rowB,colB) of the
 *  first table — the shape Merge Cells receives from the user. */
const selectCells = (
  editor: EditorType,
  rowA: number,
  colA: number,
  rowB: number,
  colB: number,
): void => {
  let from = -1;
  let to = -1;
  editor.state.doc.descendants((node, nodePos) => {
    if (node.type.name !== "table") return true;
    const at = (r: number, c: number): number => {
      let p = nodePos + 1;
      for (let i = 0; i < r; i += 1) p += node.child(i).nodeSize;
      const rowNode = node.child(r);
      for (let i = 0; i < c; i += 1) p += rowNode.child(i).nodeSize;
      return p;
    };
    from = at(rowA, colA) + 1;
    to = at(rowB, colB) + node.child(rowB).child(colB).nodeSize - 1;
    return false;
  });
  editor.commands.setTextSelection({ from, to });
};

/** Replace the document with a 2-column table carrying an explicit grid —
 *  the shape the Cell Size commands need (insert-table stamps no widths). */
const gridTable = (editor: EditorType, widths: number[], text: string[] = []): void => {
  const cell = (t: string | undefined): unknown => ({
    type: "tableCell",
    content: [
      t ? { type: "paragraph", content: [{ type: "text", text: t }] } : { type: "paragraph" },
    ],
  });
  const row = (cells: unknown[]): unknown => ({ type: "tableRow", content: cells });
  editor.commands.setContent({
    type: "doc",
    content: [
      {
        type: "table",
        attrs: { columnWidths: widths },
        content: [
          row(text.length ? text.map((t) => cell(t)) : [cell(undefined), cell(undefined)]),
          row([cell(undefined), cell(undefined)]),
        ],
      },
    ],
  } as never);
  caretInCell(editor, 0, 0);
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

  it("insert-row and insert-column create empty cells rather than duplicating cell text", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("Sample Text"));
      return true;
    });
    expect(tablesOf(editor)[0]!.child(0).child(0).textContent).toBe("Sample Text");
    expect(editor.commands["insert-row-below"]()).toBe(true);
    const tableAfterRow = tablesOf(editor)[0]!;
    expect(tableAfterRow.childCount).toBe(4);
    // Row 0 has the text, but the newly inserted row 1 must have an empty cell
    expect(tableAfterRow.child(0).child(0).textContent).toBe("Sample Text");
    expect(tableAfterRow.child(1).child(0).textContent).toBe("");
    // Column insert test:
    caretInCell(editor, 0, 0);
    expect(editor.commands["insert-column-right"]()).toBe(true);
    const tableAfterCol = tablesOf(editor)[0]!;
    expect(tableAfterCol.child(0).child(1).textContent).toBe("");
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
    expect(editor.commands["align-cell"]()).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.verticalAlign).toBe("center");
    expect(firstNodeOf(editor, "tableCell").child(0).attrs.alignment).toBe("center");
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

  it("table-style applies a preset's borders and conditional fills", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["table-style"]("light-list")).toBe(true);
    const borders = tablesOf(editor)[0]!.attrs.borders as Record<string, unknown>;
    expect(borders.top).toEqual(GRID);
    // Light List rules off the inside verticals — absent, not "none".
    expect(borders.insideVertical).toBeUndefined();
    // The header row's cells carry the preset's conditional fill (the first
    // cell in document order sits in that row).
    expect(firstNodeOf(editor, "tableCell").attrs.shading).toEqual({
      fill: "8EAADB",
      type: "clear",
    });
    // Switching presets rewrites every cell's shading (no stale bands).
    expect(editor.commands["table-style"]("table-grid")).toBe(true);
    expect(firstNodeOf(editor, "tableCell").attrs.shading).toBeNull();
    expect((tablesOf(editor)[0]!.attrs.borders as Record<string, unknown>).insideVertical).toEqual(
      GRID,
    );
    expect(editor.commands["table-style"]("bogus")).toBe(false);
  });

  it("toggle-table-look flips one tblLook flag at a time", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["toggle-table-look"]("bandRow")).toBe(true);
    expect((tablesOf(editor)[0]!.attrs.tableLook as Record<string, unknown>).bandRow).toBe(true);
    expect(editor.commands["toggle-table-look"]("bandRow")).toBe(true);
    expect((tablesOf(editor)[0]!.attrs.tableLook as Record<string, unknown>).bandRow).toBe(false);
    // Unknown flags decline.
    expect(editor.commands["toggle-table-look"]("bogus")).toBe(false);
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
  // Word's row/column picks are cell selections — every covered cell whole
  // (the same model a bar-arrow click or a cross-cell drag produces).
  it("select-table-column selects the caret's column across all rows", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 1); // middle cell of the header row
    expect(editor.commands["select-table-column"]()).toBe(true);
    const sel = editor.state.selection as unknown as {
      forEachCell(f: (node: { textContent: string }) => void): void;
    };
    const texts: string[] = [];
    sel.forEachCell((node) => texts.push(node.textContent));
    expect(texts).toHaveLength(3);
  });

  it("select-table-row selects the caret's whole row", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 1, 2); // last cell of the middle row
    expect(editor.commands["select-table-row"]()).toBe(true);
    const sel = editor.state.selection as unknown as {
      forEachCell(f: (node: { textContent: string }) => void): void;
    };
    const texts: string[] = [];
    sel.forEachCell((node) => texts.push(node.textContent));
    expect(texts).toHaveLength(3);
    // Outside a table both decline.
    editor.commands.setContent({ type: "doc", content: [{ type: "paragraph" }] });
    editor.commands.setTextSelection(2);
    expect(editor.commands["select-table-row"]()).toBe(false);
    expect(editor.commands["select-table-column"]()).toBe(false);
  });

  it("select-table selects every cell (not a NodeSelection — Backspace must empty cells, not erase the table)", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 2, 2);
    expect(editor.commands["select-table"]()).toBe(true);
    const sel = editor.state.selection as unknown as {
      forEachCell(f: (node: { textContent: string }) => void): void;
    };
    const texts: string[] = [];
    sel.forEachCell((node) => texts.push(node.textContent));
    expect(texts).toHaveLength(9);
    // Word's cell delete: the grid survives, every cell empties.
    editor.commands.command(({ tr, dispatch }) => {
      if (dispatch) (sel as unknown as { replace: (tr: unknown) => void }).replace(tr);
      return true;
    });
    let tables = 0;
    let filled = 0;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "table") {
        tables += 1;
        return false;
      }
      if (node.type.name === "tableCell" && node.textContent) filled += 1;
      return true;
    });
    expect(tables).toBe(1);
    expect(filled).toBe(0);
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

describe("merge / split table commands", () => {
  it("merge-cells folds same-row cells into a columnSpan", () => {
    const editor = build();
    editor.commands["insert-table"]();
    selectCells(editor, 0, 0, 0, 1);
    expect(editor.commands["merge-cells"]()).toBe(true);
    const row = tablesOf(editor)[0]!.child(0);
    expect(row.childCount).toBe(2);
    expect(row.child(0).attrs.columnSpan).toBe(2);
    // A second selection must hit the same table for the merge to apply.
    selectCells(editor, 0, 0, 0, 0);
    expect(editor.commands["merge-cells"]()).toBe(false);
  });

  it("merge-cells spans rows with verticalMerge restart/continue", () => {
    const editor = build();
    editor.commands["insert-table"]();
    selectCells(editor, 0, 0, 1, 1);
    expect(editor.commands["merge-cells"]()).toBe(true);
    const table = tablesOf(editor)[0]!;
    // Each spanned row folds to one cell; the first row is the restart (no
    // vMerge marker), the row below carries "continue".
    expect(table.child(0).childCount).toBe(2);
    expect(table.child(0).child(0).attrs.columnSpan).toBe(2);
    expect(table.child(0).child(0).attrs.verticalMerge).toBeFalsy();
    expect(table.child(1).childCount).toBe(2);
    expect(table.child(1).child(0).attrs.columnSpan).toBe(2);
    expect(table.child(1).child(0).attrs.verticalMerge).toBe("continue");
  });

  it("split-cell restores a merged cell to its own grid slot", () => {
    const editor = build();
    editor.commands["insert-table"]();
    selectCells(editor, 0, 0, 0, 1);
    editor.commands["merge-cells"]();
    caretInCell(editor, 0, 0);
    expect(editor.commands["split-cell"]()).toBe(true);
    const row = tablesOf(editor)[0]!.child(0);
    expect(row.childCount).toBe(3);
    expect(row.child(0).attrs.columnSpan).toBeFalsy();
    // Splitting an unmerged cell declines.
    expect(editor.commands["split-cell"]()).toBe(false);
  });

  it("split-table splits at the caret's row keeping both tables' attrs", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 1, 0);
    expect(editor.commands["split-table"]()).toBe(true);
    const tables = tablesOf(editor);
    expect(tables).toHaveLength(2);
    expect(tables[0]!.childCount).toBe(1);
    expect(tables[1]!.childCount).toBe(2);
    // The separator paragraph sits between the two tables (keeps them separate in Word).
    expect(editor.state.doc.child(1).type.name).toBe("paragraph");
    // The caret lands in the second table's first cell.
    const $from = editor.state.doc.resolve(editor.state.selection.from);
    expect($from.node(3).type.name).toBe("tableCell");
    expect($from.before(1)).toBe(
      editor.state.doc.firstChild!.nodeSize + editor.state.doc.child(1).nodeSize,
    );
    // Splitting at the first row declines.
    caretInCell(editor, 0, 0);
    expect(editor.commands["split-table"]()).toBe(false);
  });

  it("column-break inside a table delegates to split-table", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 1, 0);
    expect(editor.commands["column-break"]()).toBe(true);
    expect(tablesOf(editor)).toHaveLength(2);
  });

  it("merge-cells supports CellSelection directly from canvas drag", () => {
    const editor = build();
    editor.commands["insert-table"]();
    caretInCell(editor, 0, 0);
    // Select the first row via CellSelection.
    editor.commands["select-table-row"]();
    expect(editor.state.selection instanceof CellSelection).toBe(true);
    expect(editor.commands["merge-cells"]()).toBe(true);
    const row = tablesOf(editor)[0]!.child(0);
    expect(row.childCount).toBe(1);
    expect(row.child(0).attrs.columnSpan).toBe(3);
  });
});

describe("cell size / autofit commands", () => {
  it("autofit-window rescales the grid to the injected total width", () => {
    const editor = build();
    gridTable(editor, [720, 1440]);
    expect(editor.commands["autofit-window"]("1440")).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([480, 960]);
    expect(editor.commands["autofit-window"]("bogus")).toBe(false);
  });

  it("autofit-contents shrinks each column to its widest cell", () => {
    const editor = build();
    gridTable(editor, [2000, 2000], ["甲乙丙", ""]);
    // 甲乙丙 ≈ 3 × 240 + slack = 840 twips; the empty column hits the floor.
    expect(editor.commands["autofit-contents"]()).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([840, 720]);
  });

  it("fixed-column-width toggles the tblLayout flag", () => {
    const editor = build();
    gridTable(editor, [1440, 1440]);
    expect(editor.commands["fixed-column-width"]()).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.layout).toBe("fixed");
    expect(editor.commands["fixed-column-width"]()).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.layout).toBeNull();
  });

  it("distribute-columns splits the grid total evenly", () => {
    const editor = build();
    gridTable(editor, [720, 1440]);
    expect(editor.commands["distribute-columns"]()).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([1080, 1080]);
  });

  it("cell-width writes the caret column's width, accepting measures", () => {
    const editor = build();
    gridTable(editor, [1440, 1440]);
    // "2cm" → 2 × 1440 / 2.54 ≈ 1134 twips.
    expect(editor.commands["cell-width"]("2cm")).toBe(true);
    expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([1134, 1440]);
    // Below Word's 0.5" floor declines (1cm ≈ 567 < 720).
    expect(editor.commands["cell-width"]("1cm")).toBe(false);
  });

  it("cell-height stamps the caret row's height, auto clears it", () => {
    const editor = build();
    gridTable(editor, [1440, 1440]);
    expect(editor.commands["cell-height"]("1cm")).toBe(true);
    expect(tablesOf(editor)[0]!.child(0).attrs.height).toEqual({ value: 567, rule: "atLeast" });
    // The combobox's "auto" entry sends the clear value.
    expect(editor.commands["cell-height"]("0")).toBe(true);
    expect(tablesOf(editor)[0]!.child(0).attrs.height).toBeNull();
  });
});

describe("arrange — floating drawings", () => {
  const FLOATING = {
    behindDocument: false,
    zIndex: 2,
    horizontalPosition: { relative: "column", offset: 0 },
    verticalPosition: { relative: "paragraph", offset: 0 },
  };

  /** A doc whose paragraphs carry one floating image and one wps shape. */
  const floatDoc = (editor: EditorType): void => {
    editor.commands.setContent({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "image",
              attrs: {
                src: "data:image/png;base64,AAAA",
                width: 100,
                height: 80,
                floating: { ...FLOATING },
              },
            },
            { type: "text", text: "甲" },
          ],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "wpsShape",
              attrs: {
                wpsShape: {
                  transformation: { width: 100, height: 80 },
                  floating: { ...FLOATING, zIndex: 1 },
                },
              },
              // The editable textbox body — the node is content:"block+".
              content: [{ type: "paragraph" }],
            },
          ],
        },
      ],
    } as never);
  };

  /** Node-select the first node of `name` (the doc carries one). */
  const selectFirstNode = (editor: EditorType, name: string): void => {
    let pos = -1;
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === name) {
        pos = nodePos;
        return false;
      }
      return true;
    });
    editor.commands.setNodeSelection(pos);
  };

  it("bring-forward / send-backward step the z-order, flooring at 0", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "image");
    expect(editor.commands["bring-forward"]()).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({ zIndex: 3 });
    expect(editor.commands["send-backward"]()).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({ zIndex: 2 });

    selectFirstNode(editor, "wpsShape");
    expect(editor.commands["send-backward"]()).toBe(true);
    expect(editor.commands["send-backward"]()).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).floating,
    ).toMatchObject({ zIndex: 0 });
    expect(editor.commands["send-backward"]()).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).floating,
    ).toMatchObject({ zIndex: 0 });
  });

  it("bring-to-front tops the band, send-to-back floors it, both no-op at the extreme", () => {
    const editor = build();
    floatDoc(editor);
    // The shape (z=1) jumps past the image (z=2); a second click holds.
    selectFirstNode(editor, "wpsShape");
    expect(editor.commands["bring-to-front"]()).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).floating,
    ).toMatchObject({ zIndex: 3 });
    expect(editor.commands["bring-to-front"]()).toBe(false);
    // Against the raised shape (z=3) the image is already the floor — a
    // second drawing (shape back at z=1) exercises the actual drop.
    const other = build();
    floatDoc(other);
    selectFirstNode(other, "image");
    expect(other.commands["send-to-back"]()).toBe(true);
    expect(firstNodeOf(other, "image").attrs.floating).toMatchObject({ zIndex: 0 });
    expect(other.commands["send-to-back"]()).toBe(false);
  });

  it("bring-to-front / send-to-back stay within the drawing's band", () => {
    const editor = build();
    floatDoc(editor);
    // Move the shape behind text — its z=1 no longer bounds the front band.
    selectFirstNode(editor, "wpsShape");
    editor.commands.wrap("behind");
    selectFirstNode(editor, "image");
    // Alone in the front band, both extremes hold at the current z.
    expect(editor.commands["send-to-back"]()).toBe(false);
    expect(editor.commands["bring-to-front"]()).toBe(false);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({ zIndex: 2 });
  });

  it("wrap stamps the wrap type; front/behind clear it and set behindDoc", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "image");
    expect(editor.commands.wrap("square")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({ wrap: { type: "square" } });
    expect(editor.commands.wrap("top-bottom")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({
      wrap: { type: "topAndBottom" },
    });
    // In Front of Text drops the wrap (wrapNone) and clears behindDoc.
    expect(editor.commands.wrap("front")).toBe(true);
    const front = firstNodeOf(editor, "image").attrs.floating as Record<string, unknown>;
    expect("wrap" in front).toBe(false);
    expect(front.behindDocument).toBe(false);
    expect(editor.commands.wrap("behind")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({ behindDocument: true });
    expect(editor.commands.wrap("bogus")).toBe(false);
  });

  it("converts an inline picture to floating on Word's column/paragraph anchor", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "image", attrs: { src: "data:,", width: 10, height: 10 } }],
          },
        ],
      },
    });
    selectFirstNode(editor, "image");
    expect(editor.commands.wrap("square")).toBe(true);
    const floating = firstNodeOf(editor, "image").attrs.floating as Record<string, unknown>;
    // positionH has no "paragraph" token (ST_RelFromH) — Word's keep-position
    // conversion anchors the column horizontally, the paragraph vertically,
    // and Square carries 0.125" side distances.
    expect(floating.horizontalPosition).toEqual({ relative: "column", offset: 0 });
    expect(floating.verticalPosition).toEqual({ relative: "paragraph", offset: 0 });
    expect(floating.wrap).toEqual({ type: "square" });
    expect(floating.margins).toEqual({ left: 114300, right: 114300 });
    expect(floating.behindDocument).toBe(false);
    // "In Line with Text" drops the floating payload (Word's back-conversion).
    expect(editor.commands.wrap("inline")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toBeFalsy();
  });

  it("rotate steps image rotation and toggles the tri-state flips", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "image");
    expect(editor.commands.rotate("right")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(90);
    expect(editor.commands.rotate("right")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(180);
    expect(editor.commands.rotate("left")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(90);
    // Tri-state: omitted → true (explicit emit) → false → true.
    expect(editor.commands.rotate("flip-h")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.flipH).toBe(true);
    expect(editor.commands.rotate("flip-h")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.flipH).toBe(false);
    expect(editor.commands.rotate("flip-v")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.flipV).toBe(true);
    expect(editor.commands.rotate("bogus")).toBe(false);
  });

  it("rotate on a shape writes the nested transformation", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "wpsShape");
    expect(editor.commands.rotate("right")).toBe(true);
    const shape = firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>;
    expect(shape.transformation).toMatchObject({ rotation: 90 });
    expect(editor.commands.rotate("flip-h")).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).transformation,
    ).toMatchObject({ flipHorizontal: true });
  });

  it("position stamps margin-relative aligns on both axes", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "image");
    expect(editor.commands.position("tl")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.floating).toMatchObject({
      horizontalPosition: { relative: "margin", align: "left" },
      verticalPosition: { relative: "margin", align: "top" },
      zIndex: 2,
    });
    expect(editor.commands.position("bogus")).toBe(false);
  });

  it("align-objects aligns horizontally within the margins", () => {
    const editor = build();
    floatDoc(editor);
    selectFirstNode(editor, "wpsShape");
    expect(editor.commands["align-objects"]("center")).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).floating,
    ).toMatchObject({ horizontalPosition: { relative: "margin", align: "center" } });
    expect(editor.commands["align-objects"]("justify")).toBe(false);
    expect(editor.commands["align-objects"]()).toBe(true);
    expect(
      (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>).floating,
    ).toMatchObject({ horizontalPosition: { relative: "margin", align: "left" } });
  });

  it("declines on a bare caret, a text range, or an inline image", () => {
    const editor = build();
    floatDoc(editor);
    editor.commands.setTextSelection(2);
    expect(editor.commands["bring-forward"]()).toBe(false);
    expect(editor.commands["send-backward"]()).toBe(false);
    expect(editor.commands.wrap("square")).toBe(false);
    expect(editor.commands.rotate("right")).toBe(false);
    expect(editor.commands.position("mc")).toBe(false);
    expect(editor.commands["align-objects"]("left")).toBe(false);
  });
});

describe("add-text — TOC level stamps", () => {
  it("stamps heading levels and clears them, keeping the style rule", () => {
    const editor = build();
    editor.commands.setContent({
      type: "doc",
      content: [
        { type: "paragraph", content: [{ type: "text", text: "甲" }] },
        { type: "paragraph", content: [{ type: "text", text: "乙" }] },
        {
          type: "paragraph",
          attrs: { style: "IntenseQuote" },
          content: [{ type: "text", text: "丙" }],
        },
      ],
    });
    editor.commands.setTextSelection({ from: 2, to: editor.state.doc.content.size - 1 });
    expect(editor.commands["add-text"]("level-2")).toBe(true);
    const headingsOf = () => {
      const paras: Record<string, unknown>[] = [];
      editor.state.doc.descendants((n) => {
        if (n.type.name === "paragraph") paras.push(n.attrs as Record<string, unknown>);
      });
      return paras;
    };
    // Every selected paragraph became Heading2; the named style yields to it.
    const paras = headingsOf();
    expect(paras[0].heading).toBe("Heading2");
    expect(paras[1].heading).toBe("Heading2");
    expect(paras[2].heading).toBe("Heading2");
    expect(paras[2].style).toBeNull();
    // "none" returns them to body text (the style stays cleared).
    expect(editor.commands["add-text"]("none")).toBe(true);
    expect(headingsOf()[0].heading).toBeNull();
    expect(editor.commands["add-text"]("bogus")).toBe(false);
  });
});

/** An offset-anchored floating picture in its own paragraph (Word's anchor
 *  run shape — the image node is inline-only). */
const buildWithFloat = (floating: object): EditorType =>
  new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "image",
              attrs: { src: "data:,", width: 10, height: 10, floating },
            },
          ],
        },
      ],
    },
  });

/** Select the document's first image (the float the helpers build). */
const selectFloat = (editor: EditorType): void => {
  let pos = -1;
  editor.state.doc.descendants((node, nodePos) => {
    if (node.type.name === "image" && pos < 0) {
      pos = nodePos;
      return false;
    }
    return true;
  });
  editor.commands.setNodeSelection(pos);
};

const floatOf = (editor: EditorType): Record<string, unknown> => {
  let floating: unknown;
  editor.state.doc.descendants((node) => {
    if (node.type.name === "image" && floating === undefined) {
      floating = (node.attrs as Record<string, unknown>).floating;
      return false;
    }
    return true;
  });
  return floating as Record<string, unknown>;
};

describe("move-drawing", () => {
  it("adds the drag delta (EMU) to the selected drawing's offsets", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(editor.commands["move-drawing"](JSON.stringify({ h: 300, v: -500 }))).toBe(true);
    const floating = floatOf(editor);
    expect((floating.horizontalPosition as Record<string, unknown>).offset).toBe(1300);
    expect((floating.verticalPosition as Record<string, unknown>).offset).toBe(1500);
    // The drag keeps the drawing selected (Word's picture stays grabbed).
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("declines with no drawing selected", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    editor.commands.setTextSelection(1);
    expect(editor.commands["move-drawing"](JSON.stringify({ h: 1, v: 1 }))).toBe(false);
  });

  it("declines an inline (non-floating) picture selection", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "image", attrs: { src: "data:,", width: 10, height: 10 } }],
          },
        ],
      },
    });
    selectFloat(editor);
    expect(editor.commands["move-drawing"](JSON.stringify({ h: 1, v: 1 }))).toBe(false);
  });

  it("declines an align-anchored float — its position is not an offset", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", align: "right" },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(editor.commands["move-drawing"](JSON.stringify({ h: 1, v: 1 }))).toBe(false);
  });

  it("declines malformed JSON", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(editor.commands["move-drawing"]("not json")).toBe(false);
    expect(editor.commands["move-drawing"]()).toBe(false);
  });
});

describe("drawing-width / drawing-height", () => {
  const FLOATING = {
    horizontalPosition: { relative: "margin", offset: 1000 },
    verticalPosition: { relative: "paragraph", offset: 2000 },
  };

  it("stamps the typed measure as the selected picture's px extent", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    // 1cm ≈ 37.8px → 38; 1440 twips (1") = 96px.
    expect(editor.commands["drawing-width"]("1cm")).toBe(true);
    expect(editor.commands["drawing-height"]("1in")).toBe(true);
    const image = firstNodeOf(editor, "image");
    expect(image.attrs.width).toBe(38);
    expect(image.attrs.height).toBe(96);
    // The resize keeps the picture selected (Word's box stays active).
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("resizes an inline picture (no floating carrier required)", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "image", attrs: { src: "data:,", width: 10, height: 10 } }],
          },
        ],
      },
    });
    selectFloat(editor);
    expect(editor.commands["drawing-width"]("120px")).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.width).toBe(120);
  });

  it("declines non-positive or malformed measures and non-image selections", () => {
    const editor = buildWithFloat(FLOATING);
    expect(editor.commands["drawing-width"]("5cm")).toBe(false);
    selectFloat(editor);
    expect(editor.commands["drawing-width"]("0")).toBe(false);
    expect(editor.commands["drawing-width"]("-3cm")).toBe(false);
    expect(editor.commands["drawing-height"]("wide")).toBe(false);
    expect(editor.commands["drawing-height"]()).toBe(false);
  });
});

describe("rotate-drawing", () => {
  const FLOATING = {
    horizontalPosition: { relative: "margin", offset: 1000 },
    verticalPosition: { relative: "paragraph", offset: 2000 },
  };

  it("adds the swept degrees to the selected floating image's rotation", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(editor.commands["rotate-drawing"](JSON.stringify(45))).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(45);
    // The sweep accumulates onto the drawing's current angle.
    expect(editor.commands["rotate-drawing"](JSON.stringify(-90))).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(-45);
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("rotates a floating wps shape through its payload's transformation", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "wpsShape",
                attrs: {
                  wpsShape: {
                    floating: FLOATING,
                    transformation: { width: 914400, height: 914400, rotation: 30 },
                  },
                },
                // The shape's editable text body (a block+) rides the node.
                content: [{ type: "paragraph" }],
              },
            ],
          },
        ],
      },
    });
    let pos = -1;
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === "wpsShape" && pos < 0) {
        pos = nodePos;
        return false;
      }
      return true;
    });
    editor.commands.setNodeSelection(pos);
    expect(editor.commands["rotate-drawing"](JSON.stringify(15))).toBe(true);
    const shape = firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>;
    expect((shape.transformation as Record<string, unknown>).rotation).toBe(45);
  });

  it("rotates an inline image through the same flat attr", () => {
    // Word spins an inline picture about its extent's center exactly like a
    // floating one — the rotation attr carries either way.
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "image", attrs: { src: "data:,", width: 10, height: 10 } }],
          },
        ],
      },
    });
    selectFloat(editor);
    expect(editor.commands["rotate-drawing"](JSON.stringify(30))).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.rotation).toBe(30);
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("declines without a floating drawing selected or a zero sweep", () => {
    const editor = buildWithFloat(FLOATING);
    expect(editor.commands["rotate-drawing"](JSON.stringify(45))).toBe(false);
    selectFloat(editor);
    expect(editor.commands["rotate-drawing"](JSON.stringify(0))).toBe(false);
    expect(editor.commands["rotate-drawing"]("nan")).toBe(false);
    expect(editor.commands["rotate-drawing"]()).toBe(false);
  });
});

describe("place-drawing", () => {
  const FLOATING = {
    horizontalPosition: { relative: "margin", offset: 1000 },
    verticalPosition: { relative: "paragraph", offset: 2000 },
  };

  it("lands an align-anchored float as page-anchored offsets", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", align: "center" },
      verticalPosition: { relative: "page", align: "top" },
    });
    selectFloat(editor);
    expect(editor.commands["place-drawing"](JSON.stringify({ h: 914400, v: 1828800 }))).toBe(true);
    const floating = floatOf(editor);
    // The painted spot IS the value: relative flips to page, the offset
    // replaces the alignment outright.
    expect(floating.horizontalPosition).toEqual({ relative: "page", offset: 914400 });
    expect(floating.verticalPosition).toEqual({ relative: "page", offset: 1828800 });
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("also accepts an offset-anchored float (the bridge's dispatch is by anchor)", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(editor.commands["place-drawing"](JSON.stringify({ h: 1, v: 2 }))).toBe(true);
    const floating = floatOf(editor);
    expect(floating.horizontalPosition).toEqual({ relative: "page", offset: 1 });
  });

  it("declines malformed or incomplete values and non-floating selections", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(editor.commands["place-drawing"]("not json")).toBe(false);
    expect(editor.commands["place-drawing"](JSON.stringify({ h: 1 }))).toBe(false);
    expect(editor.commands["place-drawing"]()).toBe(false);
    // Without the drawing selected there is nothing to place.
    editor.commands.setTextSelection(1);
    expect(editor.commands["place-drawing"](JSON.stringify({ h: 1, v: 2 }))).toBe(false);
  });
});

describe("drawing-properties-apply", () => {
  const FLOATING = {
    horizontalPosition: { relative: "margin", offset: 4572000 },
    verticalPosition: { relative: "paragraph", offset: 95250 },
  };

  it("stamps the dialog's cm geometry onto the selected image", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(
      editor.commands["drawing-properties-apply"]({
        widthCm: 5,
        heightCm: 3,
        rotationDeg: 45,
        offsetHCm: 2,
        offsetVCm: 1,
      }),
    ).toBe(true);
    const attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    // cm → px at 96 DPI (5cm ≈ 188.98 → 189; 3cm ≈ 112.94 → 113).
    expect(attrs.width).toBe(189);
    expect(attrs.height).toBe(113);
    expect(attrs.rotation).toBe(45);
    // cm → EMU (360000/cm) — 2cm and 1cm exactly.
    const floating = attrs.floating as Record<string, unknown>;
    expect((floating.horizontalPosition as Record<string, unknown>).offset).toBe(720000);
    expect((floating.verticalPosition as Record<string, unknown>).offset).toBe(360000);
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("keeps the values the patch omits", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(editor.commands["drawing-properties-apply"]({ rotationDeg: 30 } as never)).toBe(true);
    const attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    expect(attrs.width).toBe(10); // untouched
    expect(attrs.rotation).toBe(30);
    const floating = attrs.floating as Record<string, unknown>;
    expect((floating.horizontalPosition as Record<string, unknown>).offset).toBe(4572000);
  });

  it("stamps and clears the alt text (the image's title attr)", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    // The command applies whatever the patch carries (the dialog sends only
    // its changed fields); the geometry values are irrelevant here — the
    // assertions read only the title.
    const patch = {
      widthCm: 0,
      heightCm: 0,
      rotationDeg: 0,
      offsetHCm: 0,
      offsetVCm: 0,
    };
    expect(editor.commands["drawing-properties-apply"]({ ...patch, altText: "XX项目效果图" })).toBe(
      true,
    );
    expect(firstNodeOf(editor, "image").attrs.title).toBe("XX项目效果图");
    // An empty field clears the text (the attr drops, defaulting to null).
    expect(editor.commands["drawing-properties-apply"]({ ...patch, altText: "" })).toBe(true);
    expect(firstNodeOf(editor, "image").attrs.title).toBeNull();
  });

  it("keeps the stored EMU offsets a partial patch omits", () => {
    // The dialog sends only the fields the user changed — an untouched
    // offset must not round-trip through the two-decimal cm display (a
    // 2000-EMU offset would land as 3600).
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    expect(editor.commands["drawing-properties-apply"]({ lockAnchor: true })).toBe(true);
    const floating = firstNodeOf(editor, "image").attrs.floating as Record<string, unknown>;
    expect((floating.horizontalPosition as Record<string, unknown>).offset).toBe(4572000);
    expect((floating.verticalPosition as Record<string, unknown>).offset).toBe(95250);
    expect(floating.lockAnchor).toBe(true);
    // Untouched wrap distances don't materialize a zero margins object.
    expect(floating.margins).toBeUndefined();
  });

  it("stamps the position bases, layout flags, and wrap distances", () => {
    const editor = buildWithFloat(FLOATING);
    selectFloat(editor);
    // The command applies whatever the patch carries; only the floating
    // extras matter to the assertions.
    const patch = {
      widthCm: 0,
      heightCm: 0,
      rotationDeg: 0,
      offsetHCm: 0,
      offsetVCm: 0,
      relativeH: "page",
      relativeV: "page",
      allowOverlap: false,
      layoutInCell: true,
      lockAnchor: true,
      distanceCm: { top: 0.13, bottom: 0.25, left: 0.13, right: 0.25 },
    };
    expect(editor.commands["drawing-properties-apply"](patch)).toBe(true);
    const floating = firstNodeOf(editor, "image").attrs.floating as Record<string, unknown>;
    expect(floating.horizontalPosition).toEqual({ relative: "page", offset: 0 });
    expect(floating.verticalPosition).toEqual({ relative: "page", offset: 0 });
    // The flags land verbatim; the distances become the EMU margins
    // (0.13cm ≈ 46800 EMU, 0.25cm = 90000).
    expect(floating.allowOverlap).toBe(false);
    expect(floating.layoutInCell).toBe(true);
    expect(floating.lockAnchor).toBe(true);
    expect(floating.margins).toEqual({ top: 46800, bottom: 90000, left: 46800, right: 90000 });
  });

  it("rides the shape payload with the same extras", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "wpsShape",
                attrs: {
                  wpsShape: {
                    transformation: { width: 100, height: 80 },
                    floating: {
                      horizontalPosition: { relative: "column", offset: 1000 },
                      verticalPosition: { relative: "paragraph", offset: 2000 },
                    },
                  },
                },
                // The editable textbox body — the node is content:"block+".
                content: [{ type: "paragraph" }],
              },
            ],
          },
        ],
      },
    });
    let shapePos = -1;
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === "wpsShape") {
        shapePos = nodePos;
        return false;
      }
      return true;
    });
    editor.commands.setNodeSelection(shapePos);
    const patch = {
      widthCm: 0,
      heightCm: 0,
      rotationDeg: 0,
      offsetHCm: 0,
      offsetVCm: 0,
      relativeH: "line",
      relativeV: "margin",
      allowOverlap: true,
      layoutInCell: false,
      lockAnchor: false,
      distanceCm: { top: 0.1, bottom: 0.1, left: 0.1, right: 0.1 },
    };
    expect(editor.commands["drawing-properties-apply"](patch)).toBe(true);
    const floating = (firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>)
      .floating as Record<string, unknown>;
    expect(floating.horizontalPosition).toEqual({ relative: "line", offset: 0 });
    expect(floating.verticalPosition).toEqual({ relative: "margin", offset: 0 });
    expect(floating.allowOverlap).toBe(true);
    expect(floating.layoutInCell).toBe(false);
    expect(floating.margins).toEqual({ top: 36000, bottom: 36000, left: 36000, right: 36000 });
  });

  it("declines without a floating drawing selected", () => {
    const editor = buildWithFloat(FLOATING);
    expect(editor.commands["drawing-properties-apply"]({ rotationDeg: 1 } as never)).toBe(false);
    expect(editor.commands["drawing-properties-apply"]()).toBe(false);
  });
});

describe("drawing-crop-apply", () => {
  it("stamps the crop fractions as the attrs' raw percentage ints", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(
      editor.commands["drawing-crop-apply"]({ left: 0.1, top: 0.25, right: 0.05, bottom: 0 }),
    ).toBe(true);
    const attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    // Fractions ×100000 = the raw ST_Percentage ints (10000 = 10%), the
    // value space cropOf reads back with a /100000.
    expect(attrs.crop).toEqual({ left: 10000, top: 25000, right: 5000, bottom: 0 });
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("resizes the extent to the kept region at the source scale", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(editor.commands["drawing-crop-apply"]({ left: 0.2, top: 0, right: 0, bottom: 0 })).toBe(
      true,
    );
    let attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    expect(attrs.width).toBe(8); // 10 × (1 − 0.2) at the unchanged scale
    expect(attrs.height).toBe(10);
    // A second crop trades kept fractions against the previous extent.
    expect(editor.commands["drawing-crop-apply"]({ left: 0.1, top: 0, right: 0, bottom: 0 })).toBe(
      true,
    );
    attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    expect(attrs.width).toBe(9); // 8 × 0.9 / 0.8
  });

  it("clears the crop on an all-zero set", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "image",
                attrs: { src: "data:image/png;base64,AAA", crop: { left: 10000 } },
              },
            ],
          },
        ],
      },
    });
    let pos = -1;
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === "image" && pos < 0) {
        pos = nodePos;
        return false;
      }
      return true;
    });
    editor.commands.setNodeSelection(pos);
    expect(editor.commands["drawing-crop-apply"]({ left: 0, top: 0, right: 0, bottom: 0 })).toBe(
      true,
    );
    const attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    // The schema's attr default reads back as null once the key is gone —
    // falsy for compile's `if (attrs.crop)` gate either way.
    expect(attrs.crop).toBeNull();
  });

  it("declines without an image selected or a malformed patch", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    expect(editor.commands["drawing-crop-apply"]({ left: 1, top: 0, right: 0, bottom: 0 })).toBe(
      false,
    );
    selectFloat(editor);
    expect(editor.commands["drawing-crop-apply"]({ left: 0.1 } as never)).toBe(false);
    expect(editor.commands["drawing-crop-apply"]()).toBe(false);
  });
});

describe("drawing-crop-reset", () => {
  it("clears the selected image's crop with the all-zero patch", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(
      editor.commands["drawing-crop-apply"]({ left: 0.1, top: 0, right: 0.05, bottom: 0 }),
    ).toBe(true);
    expect(editor.commands["drawing-crop-reset"]()).toBe(true);
    // The schema's attr default reads back as null once the key is gone.
    expect(firstNodeOf(editor, "image").attrs.crop).toBeNull();
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("grows the extent back to the full source", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    selectFloat(editor);
    expect(editor.commands["drawing-crop-apply"]({ left: 0, top: 0, right: 0, bottom: 0.5 })).toBe(
      true,
    );
    expect((firstNodeOf(editor, "image").attrs as Record<string, unknown>).height).toBe(5);
    expect(editor.commands["drawing-crop-reset"]()).toBe(true);
    const attrs = firstNodeOf(editor, "image").attrs as Record<string, unknown>;
    expect(attrs.height).toBe(10);
    expect(attrs.width).toBe(10);
  });

  it("declines without an image selected", () => {
    const editor = buildWithFloat({
      horizontalPosition: { relative: "margin", offset: 1000 },
      verticalPosition: { relative: "paragraph", offset: 2000 },
    });
    expect(editor.commands["drawing-crop-reset"]()).toBe(false);
  });
});

describe("listLevelStepPatch", () => {
  it("steps bullet level within 0-8 bounds", () => {
    expect(listLevelStepPatch({ bullet: { level: 2 } }, -1)).toEqual({ bullet: { level: 1 } });
    expect(listLevelStepPatch({ bullet: { level: 0 } }, -1)).toEqual({ bullet: { level: 0 } });
    expect(listLevelStepPatch({ bullet: { level: 8 } }, 1)).toEqual({ bullet: { level: 8 } });
  });

  it("steps numbering level keeping the reference", () => {
    expect(listLevelStepPatch({ numbering: { reference: "list-1", level: 1 } }, -1)).toEqual({
      numbering: { reference: "list-1", level: 0 },
    });
    expect(listLevelStepPatch({ numbering: { reference: "list-1", level: 0 } }, 1)).toEqual({
      numbering: { reference: "list-1", level: 1 },
    });
  });

  it("returns null for non-list paragraphs", () => {
    expect(listLevelStepPatch({}, 1)).toBeNull();
    expect(listLevelStepPatch({ style: "Normal" }, -1)).toBeNull();
  });
});

describe("paragraph alignment commands", () => {
  it("aligns paragraphs in regular text selection", () => {
    const editor = build();
    editor.commands.setContent({
      type: "doc",
      content: [
        { type: "paragraph", content: [{ type: "text", text: "Line 1" }] },
        { type: "paragraph", content: [{ type: "text", text: "Line 2" }] },
      ],
    } as never);
    editor.commands.setTextSelection(2);
    expect(editor.commands["align-center"]()).toBe(true);
    expect(editor.state.doc.child(0).attrs.alignment).toBe("center");
    expect(editor.state.doc.child(1).attrs.alignment).toBeNull();

    expect(editor.commands["align-right"]()).toBe(true);
    expect(editor.state.doc.child(0).attrs.alignment).toBe("right");

    expect(editor.commands["justify"]()).toBe(true);
    expect(editor.state.doc.child(0).attrs.alignment).toBe("both");

    expect(editor.commands["justify-distribute"]()).toBe(true);
    expect(editor.state.doc.child(0).attrs.alignment).toBe("distribute");

    expect(editor.commands["align-left"]()).toBe(true);
    expect(editor.state.doc.child(0).attrs.alignment).toBe("left");
  });

  it("aligns only selected cells when CellSelection is active", () => {
    const editor = build();
    editor.commands["insert-table"](); // 2 rows x 3 cols
    // Select column 1 (row 0 col 1 and row 1 col 1)
    caretInCell(editor, 0, 1);
    editor.commands["select-table-column"]();
    expect(editor.state.selection instanceof CellSelection).toBe(true);

    expect(editor.commands["align-center"]()).toBe(true);

    const table = tablesOf(editor)[0]!;
    // Col 1 of row 0 and row 1 should be centered:
    expect(table.child(0).child(1).child(0).attrs.alignment).toBe("center");
    expect(table.child(1).child(1).child(0).attrs.alignment).toBe("center");

    // Other columns should remain untouched:
    expect(table.child(0).child(0).child(0).attrs.alignment).toBeNull();
    expect(table.child(0).child(2).child(0).attrs.alignment).toBeNull();
    expect(table.child(1).child(0).child(0).attrs.alignment).toBeNull();
    expect(table.child(1).child(2).child(0).attrs.alignment).toBeNull();
  });
});

describe("normalizeParagraphAlignment", () => {
  it("maps end to right, start to left, and justify to both", () => {
    expect(normalizeParagraphAlignment("end")).toBe("right");
    expect(normalizeParagraphAlignment("start")).toBe("left");
    expect(normalizeParagraphAlignment("justify")).toBe("both");
    expect(normalizeParagraphAlignment("center")).toBe("center");
    expect(normalizeParagraphAlignment("distribute")).toBe("distribute");
  });

  it("defaults unknown or missing values to left", () => {
    expect(normalizeParagraphAlignment(undefined)).toBe("left");
    expect(normalizeParagraphAlignment(null)).toBe("left");
    expect(normalizeParagraphAlignment("invalid")).toBe("left");
  });
});

describe("drawing-group / drawing-ungroup", () => {
  const EMU = 9525;

  /** Two floating images at known page boxes: (200,100) 120×90 and (400,160)
   *  60×60 — union (200,100) 260×120, the first image at the union's origin. */
  const buildTwoFloats = (): EditorType =>
    new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "image",
                attrs: {
                  src: "data:,",
                  width: 120,
                  height: 90,
                  floating: {
                    horizontalPosition: { relative: "page", offset: 200 * EMU },
                    verticalPosition: { relative: "page", offset: 100 * EMU },
                  },
                },
              },
              {
                type: "image",
                attrs: {
                  src: "data:,",
                  width: 60,
                  height: 60,
                  floating: {
                    horizontalPosition: { relative: "page", offset: 400 * EMU },
                    verticalPosition: { relative: "page", offset: 160 * EMU },
                  },
                },
              },
            ],
          },
        ],
      },
    });

  /** Document positions and page boxes of every floating image, in order. */
  const floatMembers = (
    editor: EditorType,
  ): { pos: number; box: { x: number; y: number; width: number; height: number } }[] => {
    const out: { pos: number; box: { x: number; y: number; width: number; height: number } }[] = [];
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name !== "image") return true;
      const floating = (node.attrs as Record<string, unknown>).floating as
        | { horizontalPosition?: { offset?: number }; verticalPosition?: { offset?: number } }
        | undefined;
      if (!floating?.horizontalPosition?.offset || !floating.verticalPosition?.offset) return true;
      out.push({
        pos,
        box: {
          x: floating.horizontalPosition.offset / EMU,
          y: floating.verticalPosition.offset / EMU,
          width: (node.attrs as Record<string, unknown>).width as number,
          height: (node.attrs as Record<string, unknown>).height as number,
        },
      });
      return true;
    });
    return out;
  };

  const payloadOf = (
    members: { pos: number; box: { x: number; y: number; width: number; height: number } }[],
  ): string => JSON.stringify({ members });

  it("groups two floats into a fresh 1:1 group at their union", () => {
    const editor = buildTwoFloats();
    const members = floatMembers(editor);
    expect(editor.commands["drawing-group"](payloadOf(members))).toBe(true);

    const group = firstNodeOf(editor, "wpgGroup");
    const g = group.attrs.wpgGroup as Record<string, unknown>;
    // The union is the group's extent, 1:1 with its own child space.
    expect(g.transformation).toEqual({ width: 260 * EMU, height: 120 * EMU });
    expect(g.childOffsetX).toBe(0);
    expect(g.childOffsetY).toBe(0);
    expect(g.childExtentWidth).toBe(260 * EMU);
    expect(g.childExtentHeight).toBe(120 * EMU);
    // The anchor (first image, at the union's origin) keeps its offsets.
    const floating = g.floating as Record<string, any>;
    expect(floating.horizontalPosition.offset).toBe(200 * EMU);
    expect(floating.verticalPosition.offset).toBe(100 * EMU);

    // Both members ride the child space; neither keeps a floating.
    const imgs: AnyNode[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === "image") {
        imgs.push(node as unknown as AnyNode);
        return false;
      }
      return true;
    });
    expect(imgs.map((n) => n.attrs.groupXfrm)).toEqual([
      { x: 0, y: 0, cx: 120 * EMU, cy: 90 * EMU },
      { x: 200 * EMU, y: 60 * EMU, cx: 60 * EMU, cy: 60 * EMU },
    ]);
    expect(imgs.every((n) => n.attrs.floating == null)).toBe(true);
    // The group lands selected (NodeSelection), ready to drag.
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("rotate on a group writes wpgGroup, not a phantom wpsShape payload", () => {
    const editor = buildTwoFloats();
    const members = floatMembers(editor);
    expect(editor.commands["drawing-group"](payloadOf(members))).toBe(true);
    // The group command lands it selected — rotate targets that selection.
    expect(editor.commands.rotate("flip-v")).toBe(true);
    const group = firstNodeOf(editor, "wpgGroup");
    expect((group.attrs.wpgGroup as Record<string, unknown>).transformation).toMatchObject({
      flipVertical: true,
    });
    expect("wpsShape" in group.attrs).toBe(false);
  });

  it("ungroup returns every member to its pre-group position (round-trip)", () => {
    const editor = buildTwoFloats();
    const before = floatMembers(editor);
    expect(editor.commands["drawing-group"](payloadOf(before))).toBe(true);
    expect(editor.commands["drawing-ungroup"]()).toBe(true);

    // The two floats are back as standalone nodes at their original offsets
    // (±1 EMU of the px rounding through the child space) and sizes.
    const after = floatMembers(editor);
    expect(after).toHaveLength(2);
    for (let i = 0; i < 2; i += 1) {
      expect(Math.abs(after[i]!.box.x * EMU - before[i]!.box.x * EMU)).toBeLessThanOrEqual(1);
      expect(Math.abs(after[i]!.box.y * EMU - before[i]!.box.y * EMU)).toBeLessThanOrEqual(1);
      expect(after[i]!.box.width).toBe(before[i]!.box.width);
      expect(after[i]!.box.height).toBe(before[i]!.box.height);
    }
    // No group survives; the first member lands selected.
    let groups = 0;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "wpgGroup") groups += 1;
    });
    expect(groups).toBe(0);
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
  });

  it("groups a floating shape through its payload and frees it back", () => {
    const editor = new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "image",
                attrs: {
                  src: "data:,",
                  width: 100,
                  height: 50,
                  floating: {
                    horizontalPosition: { relative: "page", offset: 0 },
                    verticalPosition: { relative: "page", offset: 0 },
                  },
                },
              },
              {
                type: "wpsShape",
                attrs: {
                  wpsShape: {
                    transformation: { width: 50 * EMU, height: 50 * EMU },
                    floating: {
                      horizontalPosition: { relative: "page", offset: 150 * EMU },
                      verticalPosition: { relative: "page", offset: 25 * EMU },
                    },
                  },
                },
                content: [{ type: "paragraph" }],
              },
            ],
          },
        ],
      },
    });
    // Page boxes: image (0,0) 100×50, shape (150,25) 50×50 — union (0,0) 200×75.
    const payload = JSON.stringify({
      members: [
        { pos: 1, box: { x: 0, y: 0, width: 100, height: 50 } },
        { pos: 2, box: { x: 150, y: 25, width: 50, height: 50 } },
      ],
    });
    expect(editor.commands["drawing-group"](payload)).toBe(true);
    const shape = firstNodeOf(editor, "wpsShape");
    const ws = shape.attrs.wpsShape as Record<string, any>;
    // Child-space offset form; no floating on the payload.
    expect(ws.transformation.offset).toEqual({ left: 150 * EMU, top: 25 * EMU });
    expect(ws.transformation.width).toBe(50 * EMU);
    expect(ws.floating).toBeUndefined();

    expect(editor.commands["drawing-ungroup"]()).toBe(true);
    const freed = firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, any>;
    expect(freed.transformation.offset).toBeUndefined();
    expect(freed.transformation.width).toBe(50 * EMU);
    expect(freed.floating.horizontalPosition.offset).toBe(150 * EMU);
    expect(freed.floating.verticalPosition.offset).toBe(25 * EMU);
  });

  it("declines a single member, a non-floating pos, and an align-anchored anchor", () => {
    const editor = buildTwoFloats();
    const members = floatMembers(editor);
    expect(editor.commands["drawing-group"](JSON.stringify({ members: [members[0]!] }))).toBe(
      false,
    );
    // A paragraph pos (0) carries no floating drawing.
    expect(
      editor.commands["drawing-group"](
        JSON.stringify({ members: [members[0]!, { pos: 0, box: members[1]!.box }] }),
      ),
    ).toBe(false);
    expect(editor.commands["drawing-group"]("not json")).toBe(false);
    expect(editor.commands["drawing-group"]()).toBe(false);

    const aligned = buildTwoFloats();
    const list = floatMembers(aligned);
    // Re-anchor the first float on an align — the union has no base to sit on.
    aligned.commands.command(({ tr, dispatch }) => {
      if (dispatch) {
        tr.setNodeMarkup(list[0]!.pos, undefined, {
          ...(aligned.state.doc.nodeAt(list[0]!.pos)!.attrs as Record<string, unknown>),
          floating: {
            horizontalPosition: { relative: "margin", align: "right" },
            verticalPosition: { relative: "page", offset: 100 * EMU },
          },
        });
      }
      return true;
    });
    expect(aligned.commands["drawing-group"](payloadOf(floatMembers(aligned)))).toBe(false);
  });

  it("declines ungroup without a floating group selected", () => {
    const editor = buildTwoFloats();
    const members = floatMembers(editor);
    editor.commands["drawing-group"](payloadOf(members));
    // Drop to a caret — no NodeSelection.
    editor.commands.setTextSelection(1);
    expect(editor.commands["drawing-ungroup"]()).toBe(false);
  });
});

describe("drawing-distribute", () => {
  const EMU = 9525;

  /** Three same-size floats at x = 100 / 250 / 700 (uneven gaps 50 and 350). */
  const buildThree = (): EditorType =>
    new Editor({
      element: null,
      extensions: EXTENSIONS,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [100, 250, 700].map((x) => ({
              type: "image",
              attrs: {
                src: "data:,",
                width: 100,
                height: 100,
                floating: {
                  horizontalPosition: { relative: "page", offset: x * EMU },
                  verticalPosition: { relative: "page", offset: 0 },
                },
              },
            })),
          },
        ],
      },
    });

  const xOf = (editor: EditorType): number[] => {
    const out: number[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === "image") {
        const f = (node.attrs as Record<string, any>).floating;
        out.push(f.horizontalPosition.offset / EMU);
      }
    });
    return out;
  };

  it("equalizes the gaps; the outer two hold still", () => {
    const editor = buildThree();
    // span (800 − 100) = 700, sizes 300 → gap (700 − 300) / 2 = 200 → 100 / 400 / 700.
    const members = xOf(editor).map((x, i) => ({
      pos: i + 1,
      box: { x, y: 0, width: 100, height: 100 },
    }));
    expect(editor.commands["drawing-distribute"]("h", JSON.stringify({ members }))).toBe(true);
    expect(xOf(editor)).toEqual([100, 400, 700]);
  });

  it("declines an unknown axis, a bad payload, or a single member", () => {
    const editor = buildThree();
    const members = xOf(editor).map((x, i) => ({
      pos: i + 1,
      box: { x, y: 0, width: 100, height: 100 },
    }));
    expect(editor.commands["drawing-distribute"]("diagonal", JSON.stringify({ members }))).toBe(
      false,
    );
    expect(editor.commands["drawing-distribute"]("h", "not json")).toBe(false);
    expect(editor.commands["drawing-distribute"]("h")).toBe(false);
    expect(
      editor.commands["drawing-distribute"]("h", JSON.stringify({ members: [members[0]!] })),
    ).toBe(false);
  });
});

describe("shape-effects / shape-text-direction", () => {
  /** A floating wps text-box shape, node-selected — the Shape Styles and
   *  Text group commands' target (they write attrs wherever the shape sits,
   *  floating being one of the places). */
  const buildWithShape = (): EditorType => {
    const editor = build();
    editor.commands.setContent({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "wpsShape",
              attrs: {
                wpsShape: {
                  geometry: "rect",
                  transformation: { width: 200000, height: 120000 },
                  floating: {
                    horizontalPosition: { relative: "column", offset: 0 },
                    verticalPosition: { relative: "paragraph", offset: 0 },
                  },
                },
              },
              // The editable textbox body — the node is content:"block+".
              content: [{ type: "paragraph" }],
            },
          ],
        },
      ],
    } as never);
    let pos = -1;
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === "wpsShape") {
        pos = nodePos;
        return false;
      }
      return true;
    });
    editor.commands.setNodeSelection(pos);
    return editor;
  };

  const shapeOf = (editor: EditorType): Record<string, unknown> =>
    firstNodeOf(editor, "wpsShape").attrs.wpsShape as Record<string, unknown>;

  it("shape-effects stamps the gallery preset outer shadow and clears it", () => {
    const editor = buildWithShape();
    expect(editor.commands["shape-effects"]("shadow-lower-right")).toBe(true);
    expect(shapeOf(editor).effects).toEqual({
      outerShadow: {
        distance: 25400,
        direction: 45,
        blurRadius: 38100,
        color: { value: "000000", transforms: { alpha: 60 } },
      },
    });
    // The shape stays selected (the menu can repeat).
    expect(editor.state.selection instanceof NodeSelection).toBe(true);
    // "none" drops the emptied effects object — no husk group behind.
    expect(editor.commands["shape-effects"]("none")).toBe(true);
    expect(shapeOf(editor).effects).toBeUndefined();
    // Clearing again declines (nothing to clear) without touching the doc.
    expect(editor.commands["shape-effects"]("none")).toBe(false);
  });

  it("shape-effects declines unknown picks and non-shape selections", () => {
    const editor = buildWithShape();
    expect(editor.commands["shape-effects"]("glow")).toBe(false);
    editor.commands.setTextSelection(1);
    expect(editor.commands["shape-effects"]("shadow-right")).toBe(false);
  });

  it("shape-text-direction stamps and clears bodyPr @vert", () => {
    const editor = buildWithShape();
    expect(editor.commands["shape-text-direction"]("vertical")).toBe(true);
    expect(shapeOf(editor).bodyProperties).toEqual({ vertical: "vertical" });
    expect(editor.commands["shape-text-direction"]("vertical270")).toBe(true);
    expect(shapeOf(editor).bodyProperties).toEqual({ vertical: "vertical270" });
    // "horizontal" is the cleared state — an emptied bodyProperties drops off.
    expect(editor.commands["shape-text-direction"]("horizontal")).toBe(true);
    expect(shapeOf(editor).bodyProperties).toBeUndefined();
    expect(editor.commands["shape-text-direction"]("horizontal")).toBe(false);
  });

  it("shape-text-direction declines unknown tokens", () => {
    const editor = buildWithShape();
    expect(editor.commands["shape-text-direction"]("eastAsianVertical")).toBe(false);
    expect(editor.commands["shape-text-direction"]()).toBe(false);
  });
});
