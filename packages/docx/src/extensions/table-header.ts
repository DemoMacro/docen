import type { JSONContent } from "@tiptap/core";

import { TableHeader as BaseTableHeader } from "./tiptap";
import { bordersFromElement, renderTableCellStyles, shadingFromElement } from "./utils";

/**
 * Table header extension with nested office-open attrs (mirrors TableCell).
 *
 * Attrs mirror TableCellPropertiesOptionsBase (shading/margins/borders/width as
 * nested objects + scalar OOXML properties) plus the inherited Tiptap structural
 * names colspan/rowspan/colwidth/align (rendered: false for the OOXML-named ones).
 * DOCX round-trip is near-identity: renderDocx/parseDocx pass OOXML-native attrs
 * through and only map the Tiptap structural names colspan/rowspan/colwidth to
 * OOXML columnSpan/rowSpan/width. CSS conversion happens solely in renderHTML via
 * utils.renderTableCellStyles (consuming nested shading/verticalAlign/noWrap).
 */

// ── DOCX serialization (near-identity) ──

/** Tiptap structural attrs expressed via OOXML structural fields, not passed through. */
const SKIP_KEYS = new Set(["colspan", "rowspan", "colwidth", "children"]);

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};

  // Tiptap structural names → OOXML structural fields.
  const colspan = attrs.colspan as number | undefined;
  if (colspan && colspan > 1) opts.columnSpan = colspan;
  const rowspan = attrs.rowspan as number | undefined;
  if (rowspan && rowspan > 1) opts.rowSpan = rowspan;
  const colwidth = attrs.colwidth as number[] | null | undefined;
  // Width spans every column the cell occupies (colwidth has one entry per
  // spanned column), so sum them — using only colwidth[0] under-sizes cells
  // with colspan > 1.
  if (colwidth && colwidth.length > 0) {
    const totalPx = colwidth.reduce((sum, w) => sum + (w || 0), 0);
    if (totalPx > 0) opts.width = { size: totalPx * 15, type: "dxa" };
  }

  // Remaining OOXML-native attrs passed through verbatim (drop nulls).
  for (const [key, value] of Object.entries(attrs)) {
    if (SKIP_KEYS.has(key)) continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

export function parseDocx(opts: Record<string, unknown>): Record<string, unknown> {
  const resolved = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};

  // OOXML structural fields → Tiptap structural names (default 1 for spans).
  if (resolved.columnSpan != null) attrs.colspan = resolved.columnSpan;
  if (resolved.rowSpan != null) attrs.rowspan = resolved.rowSpan;
  if (resolved.width) {
    const twips = (resolved.width as { size: number }).size ?? 0;
    if (twips) attrs.colwidth = [Math.round(twips / 15)];
  }

  // Remaining OOXML-native opts passed through (skip structural/semantic keys).
  for (const [key, value] of Object.entries(resolved)) {
    if (key === "columnSpan" || key === "rowSpan" || key === "width" || key === "children")
      continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Attr that stores an office-open native value (not parsed from HTML) ──

const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

// ── Extension ──

export const TableHeader = BaseTableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // Nested office-open objects (parsed from HTML where CSS exists)
      shading: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => shadingFromElement(el),
      },
      borders: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => bordersFromElement(el),
      },
      verticalAlign: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
      },

      // Scalar OOXML cell properties (stored verbatim; no CSS equivalent)
      textDirection: attrNative(),
      width: attrNative(),
      margins: attrNative(),
      noWrap: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => (el.style.whiteSpace === "nowrap" ? true : null),
      },
      verticalMerge: attrNative(),
      horizontalMerge: attrNative(),
      fitText: attrNative(),
      hideMark: attrNative(),
      headers: attrNative(),
      cnfStyle: attrNative(),
    };
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
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["th", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
