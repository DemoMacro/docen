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

/** Every cell of the selection rectangle (Word's block): rows between the
 *  anchor and head rows, across the anchor..head cell columns. Rows shorter
 *  than the rectangle (post-merge) contribute what they have. */
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
  const colFrom = Math.min(ra.cell, rh.cell);
  const colTo = Math.max(ra.cell, rh.cell);
  const table = $a.node($a.depth - 1);
  const tableStart = $a.before($a.depth - 1);
  for (let r = rowFrom; r <= rowTo; r += 1) {
    const rowNode = table.child(r);
    let rowPos = tableStart + 1;
    for (let i = 0; i < r; i += 1) rowPos += table.child(i).nodeSize;
    let cellPos = rowPos + 1;
    for (let c = 0; c <= Math.min(colTo, rowNode.childCount - 1); c += 1) {
      const node = rowNode.child(c);
      if (c >= colFrom) visit(node, cellPos);
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

  /** A cell selection's delete empties the selected cells (to one
   *  blank paragraph each), the grid itself survives, preserving cell attrs. */
  replace(tr: Transaction, _content: Slice = Slice.empty): void {
    const { nodes } = tr.doc.type.schema;
    if (!nodes.tableCell) return;
    const visited: { pos: number; size: number; attrs: Record<string, unknown> }[] = [];
    this.forEachCell((node, pos) =>
      visited.push({ pos, size: node.nodeSize, attrs: node.attrs as Record<string, unknown> }),
    );
    for (let i = visited.length - 1; i >= 0; i -= 1) {
      const { pos, size, attrs } = visited[i]!;
      const emptyCell = nodes.tableCell.createAndFill(attrs);
      if (emptyCell) tr.replaceWith(pos, pos + size, emptyCell);
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

  /** The single column the cell sits in, across every row of its table. */
  static colSelection($cell: ResolvedPos): CellSelection {
    const cellIndex = $cell.index($cell.depth);
    const table = $cell.node($cell.depth - 1);
    let tablePos = $cell.before($cell.depth - 1) + 1;
    let first = -1;
    let last = -1;
    for (let r = 0; r < table.childCount; r += 1) {
      const rowNode = table.child(r);
      let cellPos = tablePos + 1;
      const c = Math.min(cellIndex, rowNode.childCount - 1);
      for (let i = 0; i < c; i += 1) cellPos += rowNode.child(i).nodeSize;
      if (first < 0) first = cellPos;
      last = cellPos;
      tablePos += rowNode.nodeSize;
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
