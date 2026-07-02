import type { TableOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Table as BaseTable } from "@tiptap/extension-table";

import {
  attrNative,
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
    node: {
      attrs: Record<string, unknown>;
      firstChild: {
        childCount: number;
        child: (i: number) => { attrs: Record<string, unknown> };
      } | null;
    };
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

    // Table width: pct → percentage; dxa/numeric → twips; "auto" or no
    // <w:tblW> → fill the page content area. Word's "auto" tables size to the
    // text column (full-width unless content is sparse); with no width the
    // browser would shrink the table to its content, so match Word with 100%.
    if (a.width && typeof a.width === "object") {
      const w = a.width as { size: number | string; type?: string };
      const numSize = typeof w.size === "string" ? parseFloat(w.size) : w.size;
      if (w.type === "pct") {
        if (typeof w.size === "string" && w.size.includes("%")) {
          // office-open keeps the percentage literal verbatim ("100%" = 100%) —
          // it is NOT fiftieths-of-a-percent, so do not divide by 50.
          styles.push(`width:${w.size}`);
        } else if (!Number.isNaN(numSize)) {
          // A bare number is fiftieths-of-a-percent (5000 = 100%) per OOXML.
          styles.push(`width:${numSize / 50}%`);
        }
      } else if (w.type === "auto") {
        styles.push("width:100%");
      } else if (numSize != null) {
        const css = twipToCss(numSize);
        if (css) styles.push(`width:${css}`);
      }
    } else {
      styles.push("width:100%");
    }
    // Cap dxa tables at the page text column. A tblW in twips can exceed the
    // content area (Word's table-width vs grid-column-sum mismatch — e.g. a
    // stray narrow trailing grid column inflates tblW past the text column),
    // which overflows the fixed page box on the right. max-width:100% shrinks
    // only over-wide tables; pct/auto/100% widths are already ≤100% so this is
    // a no-op for them.
    styles.push("max-width:100%");

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

    // colgroup columns. Collect BOTH the first-row cell colwidths (DOCX
    // <w:tcW>) and the tblGrid (columnWidths, from <w:tblGrid>).
    const firstRow = node.firstChild;
    const cellPx: number[] = [];
    if (firstRow) {
      for (let i = 0; i < firstRow.childCount; i++) {
        const cw = firstRow.child(i).attrs.colwidth as number[] | null | undefined;
        if (Array.isArray(cw) && cw.length) for (const w of cw) cellPx.push(w || 0);
        else cellPx.push(0);
      }
    }
    const hasCellWidths = cellPx.some((w) => w > 0);
    // Prefer tblGrid for the colgroup: it is the table's real column structure
    // and is IDENTICAL across every split slice, so the colgroup stays stable
    // as the paginator re-splits the table. A slice's firstRow is a mid-table
    // row whose tcW colwidths need not match the grid (and are often a 61px
    // placeholder), and ProseMirror reuses the table DOM across re-flows so a
    // firstRow-based colgroup drifts — a 0-width column then collapses its text
    // into one giant over-tall row that overflows the page. Fall back to cell
    // colwidths only when tblGrid is absent or all-zero.
    const tblGridPx = ((a.columnWidths as Array<number> | null) ?? []).map((w) =>
      Math.round((w || 0) / 15),
    );
    const hasGrid = tblGridPx.some((w) => w > 0);
    const gridPx = hasGrid ? tblGridPx : hasCellWidths ? cellPx : tblGridPx;

    if (styles.length > 0) attrs.style = styles.join(";");
    if (gridPx.some((w) => w > 0)) {
      // Column widths are relative ratios in OOXML: Word scales the grid to
      // the table width (tblW) or the page text column, never to the raw grid
      // sum. Emit percentages so the table's CSS width (tblW → pt, or 100% for
      // "auto") sets the total and columns share it proportionally. Absolute
      // px would let an oversized grid (e.g. 1063px of columns on a 579px
      // tblW) blow past the page content area, which the fixed page box then
      // clips — Word never does.
      const gridTotal = gridPx.reduce((sum, w) => sum + w, 0);
      const cols = gridPx.map(
        (w) =>
          [
            "col",
            { style: `width:${gridTotal > 0 ? ((w / gridTotal) * 100).toFixed(2) : 0}%` },
          ] as const,
      );
      // Wrap the content hole in <tbody>: ProseMirror requires a content hole
      // (0) to be the SOLE child of its parent, so it can't sit beside colgroup.
      return ["table", attrs, ["colgroup", {}, ...cols], ["tbody", 0]] as const;
    }
    return ["table", attrs, 0] as const;
  },

  renderDocx,
  parseDocx,
});
