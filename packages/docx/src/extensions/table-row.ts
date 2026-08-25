import type { TableRowOptions, TableRowPropertiesOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Node } from "@tiptap/core";

import { attrNative, cssToTwip, type DocxAttrSpec } from "./utils";

/**
 * Table row node with nested office-open attrs — fully custom (no upstream
 * table extension).
 *
 * Attrs mirror TableRowPropertiesOptionsBase (cantSplit/tableHeader/hidden as
 * booleans; height as a nested { value, rule } object; cellSpacing as a native
 * value; widthBefore/widthAfter as TableWidthProperties; etc.). DOCX round-trip
 * is near-identity: renderDocx/parseDocx pass attrs through (omitting the
 * `cells` structural key that DocxManager owns). CSS conversion happens only
 * in renderHTML.
 */

// ── DOCX serialization (near-identity: attrs mirror TableRowPropertiesOptionsBase) ──

/** Structural keys filled by DocxManager (compileTableNode). */
const SKIP_KEYS = new Set(["cells"]);

export function renderDocx(node: JSONContent): Partial<TableRowPropertiesOptions> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (SKIP_KEYS.has(key)) continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

export function parseDocx(opts: TableRowOptions): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(opts)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Extension ──

/** The attr key set the tableRow node declares — the full
 *  TableRowPropertiesOptions (Base + row-level track-change markers). The
 *  satisfies guard turns keyof drift into a compile error (same contract as
 *  docxParagraphAttrs in utils.ts). */
const docxTableRowAttrs = {
  // height: nested { value (twips), rule (HeightRule) } — parsed from CSS where possible
  height: {
    default: null,
    rendered: false,
    parseHTML: (el: HTMLElement) => {
      const css = el.style.height || el.getAttribute("height") || "";
      const twips = cssToTwip(css);
      return twips != null ? { value: twips, rule: "atLeast" } : null;
    },
  },

  // Scalar OOXML row properties (stored verbatim; no CSS equivalent)
  cantSplit: attrNative(),
  tableHeader: attrNative(),
  hidden: attrNative(),
  divId: attrNative(),
  // NOTE gridAfter/widthAfter: declared for round-trip fidelity, but
  // resolveTable deletes them when it rebuilds the trailing placeholder cells
  // they describe (an equivalent rewrite — see table.ts resolveTable).
  gridBefore: attrNative(),
  gridAfter: attrNative(),
  rowAlignment: attrNative(),
  cnfStyle: attrNative(),
  cellSpacing: attrNative(),
  widthBefore: attrNative(),
  widthAfter: attrNative(),
  // Row-level track-change markers (w:ins/w:del/w:trPrChange).
  insertion: attrNative(),
  deletion: attrNative(),
  revision: attrNative(),
} satisfies Record<keyof TableRowPropertiesOptions, DocxAttrSpec>;

export const TableRow = Node.create({
  name: "tableRow",
  content: "tableCell+",

  addAttributes() {
    return docxTableRowAttrs;
  },

  parseHTML() {
    return [{ tag: "tr" }];
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const attrs = { ...HTMLAttributes };
    const styles: string[] = [];

    // height.value is in twips → pt
    if (node.attrs.height && typeof node.attrs.height === "object") {
      const h = node.attrs.height as { value?: number };
      if (h.value != null) styles.push(`height:${h.value / 20}pt`);
    }

    if (styles.length > 0) attrs.style = styles.join(";");
    return ["tr", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
