import type { TableOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { Table as BaseTable } from "./tiptap";
import {
  alignmentFromElement,
  alignmentToCss,
  bordersFromElement,
  shadingFromElement,
  shadingToCss,
  twipToCss,
  renderBorderCSS,
} from "./utils";

/**
 * Table extension with nested office-open attrs.
 *
 * Attrs mirror TableOptions (width/float/layout/borders/alignment/margins/indent/
 * cellSpacing/tableLook/columnWidths/etc.). DOCX round-trip is near-identity:
 * renderDocx/parseDocx pass attrs through (omitting only `rows`, which DocxManager
 * rebuilds from the row/cell nodes). CSS conversion happens only in renderHTML.
 */

// ── DOCX serialization (near-identity: attrs mirror TableOptions minus rows) ──

/** Structural key rebuilt by DocxManager (compileTableNode walks the row nodes). */
const SKIP_KEYS = new Set(["rows", "columnWidthsRevision"]);

export function renderDocx(node: JSONContent): Partial<TableOptions> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (SKIP_KEYS.has(key)) continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

export function parseDocx(opts: Record<string, unknown>): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(opts)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Attr that stores an office-open native value (not parsed from HTML) ──

const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

// ── Extension ──

export const Table = BaseTable.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // Nested office-open objects (parsed from HTML where CSS exists)
      width: attrNative(),
      // tblGrid (<w:tblGrid>) — exact twips per column; kept on the table so DOCX
      // round-trips losslessly instead of being split into per-cell colwidth.
      columnWidths: attrNative(),
      indent: attrNative(),
      margins: attrNative(),
      float: attrNative(),
      borders: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => bordersFromElement(el),
      },
      shading: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => shadingFromElement(el),
      },

      // Scalar OOXML table properties
      alignment: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => alignmentFromElement(el),
      },
      layout: {
        default: null,
        rendered: false,
        parseHTML: (el: HTMLElement) => (el.style.tableLayout === "fixed" ? "fixed" : null),
      },
      style: attrNative(),
      visuallyRightToLeft: attrNative(),
      tableLook: attrNative(),
      cellSpacing: attrNative(),
      styleRowBandSize: attrNative(),
      styleColBandSize: attrNative(),
      caption: attrNative(),
      description: attrNative(),
    };
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const a = node.attrs;
    const attrs = { ...HTMLAttributes };
    const styles: string[] = [];

    const align = alignmentToCss(a.alignment as string | null | undefined);
    if (align === "center") {
      styles.push("margin-left:auto", "margin-right:auto");
    } else if (align === "right") {
      styles.push("margin-left:auto", "margin-right:0");
    } else if (align) {
      styles.push("margin-left:0", "margin-right:auto");
    }

    if (a.layout === "fixed") styles.push("table-layout:fixed");

    if (a.width && typeof a.width === "object") {
      const w = a.width as { size: number | string; type?: string };
      // pct size is fiftieths-of-a-percent (5000 = 100%); @office-open/docx may
      // emit it as a "5000%" string — parse to a number first.
      const numSize = typeof w.size === "string" ? parseFloat(w.size) : w.size;
      if (w.type === "pct") {
        if (!Number.isNaN(numSize)) styles.push(`width:${numSize / 50}%`);
      } else if (numSize != null) {
        const css = twipToCss(numSize);
        if (css) styles.push(`width:${css}`);
      }
    }

    if (a.indent && typeof a.indent === "object") {
      const ind = a.indent as { size?: number; type?: string };
      if (ind.size != null) {
        const css = twipToCss(ind.size);
        if (css) styles.push(`margin-left:${css}`);
      }
    }

    if (a.cellSpacing && typeof a.cellSpacing === "object") {
      // TableCellSpacingProperties: { value, type }
      const cs = a.cellSpacing as { value?: number };
      if (cs.value != null) {
        const css = twipToCss(cs.value);
        if (css) styles.push(`border-spacing:${css}`);
      }
    }

    const bg = shadingToCss(a.shading as { fill?: string } | null | undefined);
    if (bg) styles.push(`background-color:${bg}`);

    if (a.borders && typeof a.borders === "object") {
      const b = a.borders as Record<string, unknown>;
      const sides: Array<[string, unknown]> = [
        ["top", b.top],
        ["bottom", b.bottom],
        ["left", b.left],
        ["right", b.right],
      ];
      for (const [side, border] of sides) {
        const css = renderBorderCSS(border as Parameters<typeof renderBorderCSS>[0]);
        if (css) styles.push(`border-${side}:${css}`);
      }
    }

    if (styles.length > 0) attrs.style = styles.join(";");
    return ["table", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
