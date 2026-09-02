// A cell selection — Word's cross-cell drag model: anchor and head cells
// bound a rectangular block of cells, every one of them selected whole.
//
// Built on the editor's own table schema (table > tableRow > tableCell with
// OOXML-style attrs — columnSpan/verticalMerge, no colspan/rowspan), so
// prosemirror-tables' TableMap cannot read it; the grid walk here is the
// schema's own row/cell sibling order, which is exactly the order the
// painter and the caret map consume.

import type { Node as PmNode, ResolvedPos } from "@tiptap/pm/model";
import { Slice } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import { Selection, TextSelection } from "@tiptap/pm/state";
import type { Mappable } from "@tiptap/pm/transform";

/** The cell enclosing (or at) a position — its start position resolved. Null
 *  outside tables. */
export function cellAt($pos: ResolvedPos): ResolvedPos | null {
  const cellType = $pos.doc.type.schema.nodes.tableCell;
  if (!cellType) return null;
  for (let d = $pos.depth; d > 0; d -= 1) {
    if ($pos.node(d).type === cellType) return $pos.doc.resolve($pos.before(d));
  }
  return null;
}

/** {@link cellAt}, forgiving at node boundaries: a mapped cell position may
 *  land on its table's or row's start (a sibling deletion), where the walk
 *  steps forward into the first cell. */
function cellNear(doc: PmNode, pos: number): ResolvedPos | null {
  const direct = cellAt(doc.resolve(pos));
  if (direct) return direct;
  const { tableCell, tableRow, table } = doc.type.schema.nodes;
  const next = doc.resolve(pos).nodeAfter;
  if (next?.type === tableCell) return doc.resolve(pos);
  if (next?.type === tableRow || next?.type === table) return cellNear(doc, pos + 1);
  return null;
}

/** True when both cell positions belong to the same table. */
export function inSameTable(a: ResolvedPos, b: ResolvedPos): boolean {
  return a.depth >= 2 && b.depth >= 2 && a.node(a.depth - 1) === b.node(b.depth - 1);
}

/** The (row, cell) sibling indexes of a cell position within its table —
 *  cellAt's resolve sits in the row layer, so depth is the row depth. */
const rowCellOf = ($cell: ResolvedPos): { row: number; cell: number } => ({
  row: $cell.index($cell.depth - 1),
  cell: $cell.index($cell.depth),
});

/** A cell's columnSpan — 1 unless the schema attr carries a wider one. */
const spanOf = (cell: PmNode): number => {
  const span = cell.attrs.columnSpan as unknown;
  return typeof span === "number" && span > 1 ? span : 1;
};

/** The grid column a row's cell sibling starts at — the spans of its
 *  preceding siblings (a spanning cell covers its columns). */
const gridColOf = (rowNode: PmNode, cellIndex: number): number => {
  let col = 0;
  for (let i = 0; i < cellIndex; i += 1) col += spanOf(rowNode.child(i)!);
  return col;
};

/** Every cell of the selection rectangle (Word's block): rows between the
 *  anchor and head rows, across the anchor..head columns. The rectangle
 *  lives in GRID columns — a spanning cell covers its columns — so a drag
 *  over merged rows still reaches every grid column its ends span; rows
 *  shorter than the rectangle contribute what they have. */
export function cellsInRect(
  doc: PmNode,
  anchorPos: number,
  headPos: number,
  visit: (node: PmNode, pos: number) => void,
): void {
  const $a = doc.resolve(anchorPos);
  const $h = doc.resolve(headPos);
  const ra = rowCellOf($a);
  const rh = rowCellOf($h);
  const rowFrom = Math.min(ra.row, rh.row);
  const rowTo = Math.max(ra.row, rh.row);
  const table = $a.node($a.depth - 1);
  const tableStart = $a.before($a.depth - 1);
  const anchorRow = table.child(ra.row)!;
  const headRow = table.child(rh.row)!;
  // Half-open [colFrom, colTo): each end covers its own span.
  const colFrom = Math.min(gridColOf(anchorRow, ra.cell), gridColOf(headRow, rh.cell));
  const colTo = Math.max(
    gridColOf(anchorRow, ra.cell) + spanOf(anchorRow.child(ra.cell)!),
    gridColOf(headRow, rh.cell) + spanOf(headRow.child(rh.cell)!),
  );
  for (let r = rowFrom; r <= rowTo; r += 1) {
    const rowNode = table.child(r);
    let rowPos = tableStart + 1;
    for (let i = 0; i < r; i += 1) rowPos += table.child(i).nodeSize;
    let cellPos = rowPos + 1;
    let col = 0;
    for (let c = 0; c < rowNode.childCount && col < colTo; c += 1) {
      const node = rowNode.child(c);
      const span = spanOf(node);
      if (col + span > colFrom) visit(node, cellPos);
      col += span;
      cellPos += node.nodeSize;
    }
  }
}

const JSON_ID = "docen-cell";

/** A bookmark pair — PM's Selection.fromJSON/bookmark contract. */
class CellBookmark {
  constructor(
    readonly anchor: number,
    readonly head: number,
  ) {}

  map(mapping: Mappable): CellBookmark {
    return new CellBookmark(mapping.map(this.anchor), mapping.map(this.head));
  }

  resolve(doc: PmNode): Selection {
    const $anchor = cellNear(doc, this.anchor);
    const $head = cellNear(doc, this.head);
    if ($anchor && $head && inSameTable($anchor, $head)) return new CellSelection($anchor, $head);
    return TextSelection.between(doc.resolve(this.anchor), doc.resolve(this.head));
  }

  getBookmark(): CellBookmark {
    return this;
  }
}

export class CellSelection extends Selection {
  /** The anchor and head cells, at their start positions. */
  readonly anchorCell: number;
  readonly headCell: number;

  constructor($anchorCell: ResolvedPos, $headCell: ResolvedPos = $anchorCell) {
    const doc = $anchorCell.doc;
    const anchor = $anchorCell.pos;
    const head = $headCell.pos;
    const anchorEnd = anchor + doc.nodeAt(anchor)!.nodeSize;
    const headEnd = head + doc.nodeAt(head)!.nodeSize;
    super(doc.resolve(Math.min(anchor, head)), doc.resolve(Math.max(anchorEnd, headEnd)));
    this.anchorCell = anchor;
    this.headCell = head;
  }

  map(doc: PmNode, mapping: Mappable): Selection {
    const $anchor = cellNear(doc, mapping.map(this.anchorCell));
    const $head = cellNear(doc, mapping.map(this.headCell));
    // Both ends landing in (possibly moved) cells keep the cell shape;
    // anything else falls back to a text selection over the mapped range.
    if ($anchor && $head && inSameTable($anchor, $head)) {
      return new CellSelection($anchor, $head);
    }
    return TextSelection.between(
      doc.resolve(mapping.map(this.from)),
      doc.resolve(mapping.map(this.to)),
    );
  }

  /** The covered slice — the rectangle's full row range. */
  content(): Slice {
    return this.$from.doc.slice(this.from, this.to);
  }

  /** Every cell of the selection rectangle, document order. */
  forEachCell(f: (node: PmNode, pos: number) => void): void {
    cellsInRect(this.$from.doc, this.anchorCell, this.headCell, f);
  }

  /** Word's cell delete: every selected cell's content empties out (one
   *  blank paragraph each), the grid itself survives. */
  replace(tr: Transaction, _content: Slice = Slice.empty): void {
    const { nodes } = tr.doc.type.schema;
    const emptyCell = nodes.tableCell?.createAndFill();
    if (!emptyCell) return;
    const visited: { pos: number; size: number }[] = [];
    this.forEachCell((node, pos) => visited.push({ pos, size: node.nodeSize }));
    for (let i = visited.length - 1; i >= 0; i -= 1) {
      const { pos, size } = visited[i]!;
      tr.replaceWith(pos, pos + size, emptyCell);
    }
    // A caret near the range's old start — replacements shifted the cells,
    // so `near` clamps into the first emptied one.
    tr.setSelection(TextSelection.near(tr.doc.resolve(this.from)));
  }

  eq(other: unknown): boolean {
    return (
      other instanceof CellSelection &&
      other.anchorCell === this.anchorCell &&
      other.headCell === this.headCell
    );
  }

  getBookmark(): CellBookmark {
    return new CellBookmark(this.anchorCell, this.headCell);
  }

  toJSON(): { type: string; anchor: number; head: number } {
    return { type: JSON_ID, anchor: this.anchorCell, head: this.headCell };
  }

  static fromJSON(doc: PmNode, json: { anchor: number; head: number }): CellSelection {
    return new CellSelection(doc.resolve(json.anchor), doc.resolve(json.head));
  }

  static create(doc: PmNode, anchorCell: number, headCell: number = anchorCell): CellSelection {
    return new CellSelection(doc.resolve(anchorCell), doc.resolve(headCell));
  }

  /** The single row the cell sits in — anchor and head stretched to the
   *  row's first and last cells, so the rectangle covers the whole row. */
  static rowSelection($cell: ResolvedPos): CellSelection {
    const rowNode = $cell.node($cell.depth);
    let cellPos = $cell.before($cell.depth) + 1;
    const positions: number[] = [];
    for (let c = 0; c < rowNode.childCount; c += 1) {
      positions.push(cellPos);
      cellPos += rowNode.child(c).nodeSize;
    }
    return new CellSelection(
      $cell.doc.resolve(positions[0]!),
      $cell.doc.resolve(positions[positions.length - 1]!),
    );
  }

  /** The single column the cell sits in, across every row of its table. The
   *  per-row hit is the cell covering the column in GRID terms — a row that
   *  merges the column away contributes its spanning cell, not whatever
   *  sibling happens to sit at the index. */
  static colSelection($cell: ResolvedPos): CellSelection {
    const rowNode = $cell.node($cell.depth);
    const gridCol = gridColOf(rowNode, $cell.index($cell.depth));
    const table = $cell.node($cell.depth - 1);
    let tablePos = $cell.before($cell.depth - 1) + 1;
    let first = -1;
    let last = -1;
    for (let r = 0; r < table.childCount; r += 1) {
      const row = table.child(r);
      let cellPos = tablePos + 1;
      let col = 0;
      for (let c = 0; c < row.childCount && col <= gridCol; c += 1) {
        const node = row.child(c);
        const span = spanOf(node);
        if (col + span > gridCol) {
          if (first < 0) first = cellPos;
          last = cellPos;
          break;
        }
        col += span;
        cellPos += node.nodeSize;
      }
      tablePos += row.nodeSize;
    }
    return new CellSelection($cell.doc.resolve(first), $cell.doc.resolve(last));
  }

  /** The whole table the cell sits in — anchor at the first cell, head at
   *  the last (Word's corner-handle pick; Backspace then empties cells
   *  instead of erasing the table node). */
  static tableSelection($cell: ResolvedPos): CellSelection {
    const table = $cell.node($cell.depth - 1);
    let tablePos = $cell.before($cell.depth - 1) + 1;
    let first = -1;
    let last = -1;
    for (let r = 0; r < table.childCount; r += 1) {
      const rowNode = table.child(r);
      let cellPos = tablePos + 1;
      for (let c = 0; c < rowNode.childCount; c += 1) {
        if (first < 0) first = cellPos;
        last = cellPos;
        cellPos += rowNode.child(c).nodeSize;
      }
      tablePos += rowNode.nodeSize;
    }
    return new CellSelection($cell.doc.resolve(first), $cell.doc.resolve(last));
  }
}

// Registers the JSON id so a serialized selection round-trips (PM's
// Selection.fromJSON dispatches on this table). Namespaced — the plain
// "cell" id belongs to prosemirror-tables, whose schema this isn't.
Selection.jsonID(JSON_ID, CellSelection);
