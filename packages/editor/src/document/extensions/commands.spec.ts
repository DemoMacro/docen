// @vitest-environment node
import { Document, Image, Paragraph, Table, TableCell, TableRow, WpsShape } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DocumentCommands } from "./commands";

// Tiptap's schema needs the plain text node (same trick as the TOC spec).
const Text = TextNode.create({ name: "text", group: "inline" });

/** A headless editor with exactly the schema the table commands touch. */
const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: [
      Document,
      Paragraph,
      Text,
      Table,
      TableRow,
      TableCell,
      Image,
      WpsShape,
      DocumentCommands,
    ],
    content: { type: "doc", content: [{ type: "paragraph" }] },
  });

type AnyNode = {
  attrs: Record<string, unknown>;
  childCount: number;
  child: (i: number) => AnyNode;
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
    // The caret lands in the second table's first cell.
    const $from = editor.state.doc.resolve(editor.state.selection.from);
    // The caret sits inside the second table (the one after the split).
    expect($from.node(3).type.name).toBe("tableCell");
    expect($from.before(1)).toBe(editor.state.doc.firstChild!.nodeSize);
    // Splitting at the first row declines.
    caretInCell(editor, 0, 0);
    expect(editor.commands["split-table"]()).toBe(false);
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
