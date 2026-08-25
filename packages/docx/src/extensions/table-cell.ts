import type { TableCellOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Node } from "@tiptap/core";

import { docxTableCellAttrs, renderTableCellStyles } from "./utils";

/**
 * Table cell node with nested office-open attrs — fully custom (no upstream
 * table extension): attrs mirror TableCellOptions directly
 * (columnSpan/verticalMerge/width and every TableCellPropertiesOptions key),
 * so the OOXML grid shape IS the PM shape. renderDocx/parseDocx are plain
 * pass-throughs (no vMerge↔rowspan rebuild, no px colwidth↔twip conversion);
 * the layout engine expands verticalMerge into rowspan at the single
 * projection point. CSS conversion happens solely in renderHTML via
 * utils.renderTableCellStyles (consuming nested shading/verticalAlign/noWrap).
 */

// ── DOCX serialization (near-identity) ──

/** Structural keys DocxManager owns; everything else passes through. */
const SKIP_KEYS = new Set(["children"]);

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (SKIP_KEYS.has(key)) continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

export function parseDocx(opts: TableCellOptions): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(opts)) {
    if (key === "rowSpan" || key === "children" || key === "text" || key === "cellProperties")
      continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Extension ──

export const TableCell = Node.create({
  name: "tableCell",
  // A cell is a SectionChild[] block stream (like a section body or a
  // header/footer slot): paragraphs, nested tables, lists — isolating so
  // selections never straddle the cell boundary.
  content: "block+",
  isolating: true,

  addAttributes() {
    return docxTableCellAttrs();
  },

  parseHTML() {
    return [
      {
        tag: "td",
        // The HTML colspan surfaces as the OOXML columnSpan. An HTML rowspan
        // (the expanded form) has no single-cell vMerge equivalent — dropped.
        getAttrs: (el) => {
          const span = Number((el as HTMLElement).getAttribute("colspan"));
          return Number.isInteger(span) && span > 1 ? { columnSpan: span } : {};
        },
      },
      // A <th> is a header cell — Word models header-ness as the ROW's
      // tblHeader, so both tags resolve to the same node.
      { tag: "th" },
    ];
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const styles = renderTableCellStyles(node.attrs);
    const attrs = { ...HTMLAttributes };
    // The OOXML column span surfaces as the HTML colspan the browser needs.
    const span = typeof node.attrs.columnSpan === "number" ? node.attrs.columnSpan : undefined;
    if (span && span > 1) attrs.colspan = String(span);
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["td", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
