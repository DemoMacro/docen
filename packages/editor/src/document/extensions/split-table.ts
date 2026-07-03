import { Table } from "@docen/docx/extensions/table";
import { TableRow } from "@docen/docx/extensions/table-row";
import { tableFloatToCss } from "@docen/docx/extensions/utils";
import { TableView } from "@tiptap/extension-table";
import type { Node as PmNode } from "@tiptap/pm/model";
import type { EditorView } from "@tiptap/pm/view";

/**
 * Table split support for C-route pagination.
 *
 * contenteditable cannot visually break a `<tr>`, so a table taller than a page
 * is split into multiple `table` nodes across pages (a physical split). All
 * splits of one original table share a `splitGroup` id; `unwrapPages` (in
 * utils/merge.ts) merges them back into one table on export, so the split is
 * editor-only and round-trip-transparent. Continuation pages repeat the header
 * rows: the paginator clones each `tableHeader` row and marks the clone
 * `splitClone`; `unwrapPages` drops clones on merge (the original header row
 * keeps `tableHeader` → DOCX `<w:tblHeader/>`, Word repeats it natively on its
 * own pages, so no physical clone is written to DOCX).
 *
 * See CLAUDE.md → Pagination Architecture (C-route) and CONTRIBUTING.md.
 */

/** TableView variant that applies a floating table's w:tblpPr anchor (page/
 *  margin anchor → position:absolute) to the wrapper div. The base TableView
 *  rebuilds the <table> from addAttributes fields and never runs the docx
 *  node-level renderHTML, so renderHTML's computed style (float/border/width)
 *  never reaches the editor view. This re-applies the float positioning — the
 *  only table-level style that pins a table to the page box — onto the wrapper
 *  div, which the base view leaves without inline style, so it never clashes
 *  with updateColumns' table.style.width. */
class FloatTableView extends TableView {
  constructor(
    node: PmNode,
    cellMinWidth: number,
    view?: EditorView,
    HTMLAttributes?: Record<string, unknown>,
  ) {
    super(node, cellMinWidth, view, HTMLAttributes);
    this.applyFloat(node);
  }

  update(node: PmNode): boolean {
    const ok = super.update(node);
    if (ok) this.applyFloat(node);
    return ok;
  }

  private applyFloat(node: PmNode): void {
    const css = tableFloatToCss(node.attrs.float);
    this.dom.style.cssText = css.length ? css.join(";") : "";
  }
}

/** Whole-table split id shared by all table nodes split from one original.
 *  null on an un-split (whole) table. Editor-only — cleared on export. */
export const SplitTable = Table.extend({
  name: "table",
  addAttributes() {
    return {
      ...this.parent?.(),
      splitGroup: { default: null, parseHTML: () => null, rendered: false },
    };
  },
  addNodeView() {
    return ({ node, view, HTMLAttributes }) =>
      new FloatTableView(node, this.options.cellMinWidth, view, HTMLAttributes);
  },
});

/** Marks a row cloned onto a continuation page as a repeated header (not the
 *  original header row). Editor-only — dropped on export merge. */
export const SplitTableRow = TableRow.extend({
  name: "tableRow",
  addAttributes() {
    return {
      ...this.parent?.(),
      splitClone: { default: false, parseHTML: () => null, rendered: false },
    };
  },
});

/** Header rows of a table: rows repeated on every continuation page. A row is
 *  a header if the docx `tableHeader` attr is true (set on DOCX import) OR its
 *  first cell is a `tableHeader` node (<th>, as Tiptap's insertTable produces)
 *  — both mark the row for DOCX `<w:tblHeader/>` repetition. */
export function headerRowsOf(table: PmNode): PmNode[] {
  const rows: PmNode[] = [];
  table.forEach((row) => {
    const isHeader =
      row.attrs.tableHeader === true ||
      (row.firstChild != null && row.firstChild.type.name === "tableHeader");
    if (isHeader) rows.push(row);
  });
  return rows;
}

/** Clone a table's header rows for a continuation page, marking each clone
 *  `splitClone` so the export merge drops it (the original header row stays). */
export function cloneHeaderRows(table: PmNode): PmNode[] {
  return headerRowsOf(table).map((row) =>
    row.type.create({ ...row.attrs, splitClone: true }, row.content, row.marks),
  );
}
