import type {
  BorderOptions,
  SectionChild,
  TableCellOptions,
  TableOptions,
} from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Table as BaseTable } from "@tiptap/extension-table";

import { allBordersNone, cleanAttrs, mergeTableStyleProps } from "../converters/styles";
import type { ParseBlockRule, ResolveContext } from "./types";
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

// ── Block parse rule (resolve: SectionChild → table node) ──

/**
 * Declarative block parse rule: recognize a table SectionChild and rebuild it
 * as a Tiptap table node (rows/cells with colspan/rowspan recovered, the table
 * style's tblBorders/tblCellMar merged in, insideH/V grid lines pushed onto
 * cells, gridAfter as trailing nil-bordered cells). DocxManager dispatches
 * every SectionChild through this rule before the paragraph/passthrough
 * fallbacks. */
export const parseDocxBlock: ParseBlockRule = {
  match: (child) => "table" in child,
  convert: (child, ctx) =>
    resolveTable((child as unknown as { table: Record<string, unknown> }).table, ctx),
};

/** Resolve a table SectionChild into a Tiptap table node. A cell is itself a
 *  SectionChild[] block stream, resolved recursively via ctx. */
function resolveTable(tableOpts: Record<string, unknown>, ctx: ResolveContext): JSONContent {
  const attrs = ctx.parseNodeAttrs("table", tableOpts);
  const rows = (tableOpts.rows ?? []) as Record<string, unknown>[];
  const content: JSONContent[] = [];

  // Pull the referenced table style's tblBorders/tblCellMar in: office-open
  // leaves table.borders/cellMargin reflecting only the table's own tblPr, so
  // a "Table Grid" table (borders defined in the style) would render no grid
  // without this. The table's own real borders win; the style fills the gap
  // when the table's are all none/nil.
  const styleProps = mergeTableStyleProps(
    ctx.styles?.tableStyles,
    (tableOpts.style as string | undefined) ?? null,
  );
  if (styleProps.borders && allBordersNone(tableOpts.borders)) {
    tableOpts = { ...tableOpts, borders: styleProps.borders };
  }
  if (styleProps.cellMargin && tableOpts.cellMargin == null && tableOpts.margins == null) {
    tableOpts = { ...tableOpts, cellMargin: styleProps.cellMargin };
  }

  // Table-level default cell insets (w:tblCellMar). office-open exposes them
  // on both `cellMargin` and `margins`; a cell inherits them unless it carries
  // its own tcMar. Push the default onto cells without tcMar so render
  // (renderTableCellStyles) and the paginator (cellVerticalOverhead) read ONE
  // effective source (cell.attrs.margins) instead of each falling back to the
  // table. compileTableCellNode drops a cell tcMar equal to this default to
  // keep the regenerated docx in its table-level form (near-identity round-trip).
  const tableCellMargins = (tableOpts.cellMargin ?? tableOpts.margins ?? null) as NonNullable<
    TableCellOptions["margins"]
  > | null;
  // Table-level inside grid lines (tblBorders.insideHorizontal/insideVertical).
  // In CSS border-collapse the interior grid belongs to cells, not the <table>
  // element, so a REAL insideH/V is pushed onto cell sides lacking their own
  // tcBorder (below). none/nil is skipped — a table that merely LACKS inner
  // grid lines must not have `border:none` stamped on every cell, or it would
  // suppress the editor's Table-Grid fallback default for borderless tables.
  // Edge cells' outer sides overlap the table's own border under
  // border-collapse (thicker wins), matching OOXML outer-vs-inner semantics.
  const tableBorders = (tableOpts.borders ?? null) as {
    insideHorizontal?: BorderOptions;
    insideVertical?: BorderOptions;
  } | null;
  const realBorder = (bd: unknown): BorderOptions | null => {
    const b = bd as BorderOptions | null | undefined;
    return b && b.style && b.style !== "none" && b.style !== "nil" ? b : null;
  };
  const insideH = realBorder(tableBorders?.insideHorizontal);
  const insideV = realBorder(tableBorders?.insideVertical);

  // DOCX encodes a row span as a `restart` cell followed by N empty `continue`
  // cells; ProseMirror uses one cell with rowspan = N+1. Track open spans per
  // column so each continuation cell increments the owning cell's `rowspan`,
  // which compileTableNode reads back to rebuild the continuation cells.
  let activeSpans = new Map<number, JSONContent>();

  for (const row of rows) {
    const rowAttrs = ctx.parseNodeAttrs("tableRow", row as unknown as Record<string, unknown>);
    // gridAfter/widthAfter are rebuilt as trailing placeholder cells below
    // (PM requires every row's cell-span sum to equal the column count, so a
    // gridAfter row needs explicit empty cells or fixTables inserts the filler
    // at the START). Drop them from rowAttrs so compile doesn't re-emit
    // row.gridAfter on top of those cells — that double-counts and the row
    // widens by N columns every docx→json→docx round-trip.
    delete rowAttrs.gridAfter;
    delete rowAttrs.widthAfter;
    const cells = (row.cells ?? []) as Record<string, unknown>[];
    const cellNodes: JSONContent[] = [];
    const nextActiveSpans = new Map<number, JSONContent>();
    let colIdx = 0;

    for (const cell of cells) {
      const cellColspan = (cell.columnSpan as number) ?? 1;
      const vMerge = cell.verticalMerge as string | undefined;

      if (vMerge === "continue") {
        // Column owned by a cell above — bump its rowspan and carry the span
        // forward so further continuation cells keep counting.
        const owner = activeSpans.get(colIdx);
        if (owner) {
          const ownerAttrs = (owner.attrs ??= {});
          ownerAttrs.rowspan = ((ownerAttrs.rowspan as number) ?? 1) + 1;
          for (let c = colIdx; c < colIdx + cellColspan; c++) nextActiveSpans.set(c, owner);
        }
        colIdx += cellColspan;
        continue;
      }

      const isHeader = row.tableHeader as boolean;
      const cellAttrs = ctx.parseNodeAttrs(
        isHeader ? "tableHeader" : "tableCell",
        cell as unknown as Record<string, unknown>,
      );

      // `rowspan` (recovered below) drives compile-time vMerge — drop the
      // OOXML marker so it doesn't round-trip back verbatim.
      delete cellAttrs.verticalMerge;

      // Effective cell margins: a cell's own tcMar wins, else inherit the
      // table's tblCellMar default (resolved once here for render + measure).
      if (!cellAttrs.margins && tableCellMargins) cellAttrs.margins = tableCellMargins;

      // Effective cell borders: a cell's own tcBorder per side wins, else
      // inherit the table's inside grid lines (insideH on top/bottom, insideV
      // on left/right) so the interior grid renders under border-collapse.
      // compileTableCellNode drops a side equal to the table's insideH/V to
      // keep the round-trip near-identity.
      if (insideH || insideV) {
        if (!cellAttrs.borders) cellAttrs.borders = {};
        const b = cellAttrs.borders as Record<string, BorderOptions | undefined>;
        if (insideH && !b.top) b.top = insideH;
        if (insideH && !b.bottom) b.bottom = insideH;
        if (insideV && !b.left) b.left = insideV;
        if (insideV && !b.right) b.right = insideV;
      }

      // A cell is just another SectionChild[] block stream — same as a
      // section body or a header/footer slot — so resolve it through the same
      // path. That regroups consecutive numbering/bullet paragraphs into list
      // nodes and keeps nested tables/lists structurally intact on import.
      const cellChildren = (cell.children ?? []) as SectionChild[];
      const cellContent: JSONContent[] = ctx.resolveBlockStream(cellChildren);

      const cellType = isHeader ? "tableHeader" : "tableCell";
      const cellNode: JSONContent = { type: cellType };
      if (Object.keys(cellAttrs).length > 0) cellNode.attrs = cleanAttrs(cellAttrs);
      // An empty cell still needs content to satisfy the tableCell/tableHeader
      // `block+` schema. A content-less cell reaches the doc via fromJSON (which
      // skips validation), but prosemirror-tables' fixTables runs setNodeMarkup
      // on every table during appendTransaction, and setNodeMarkup re-validates
      // the node content — throwing "Invalid content for node type tableCell".
      // That throw aborts the paginator's reflow transaction, so the document
      // never re-pages (every block piles on page 0). Backfill an empty
      // paragraph; OOXML likewise requires a <w:p> in every <w:tc>.
      if (cellContent.length > 0) cellNode.content = cellContent;
      else cellNode.content = [{ type: "paragraph" }];

      if (vMerge === "restart") {
        // rowspan is finalized when continuation cells arrive below; register
        // the node so they can find and increment it.
        for (let c = colIdx; c < colIdx + cellColspan; c++) nextActiveSpans.set(c, cellNode);
      }

      cellNodes.push(cellNode);
      colIdx += cellColspan;
    }

    // OOXML gridAfter (w:gridAfter + widthAfter): N trailing grid columns this
    // row leaves uncovered. ProseMirror requires every row's cell-span sum to
    // equal the column count; without explicit trailing cells fixTables fills
    // the gap — and for a row whose only real cell is a leading gridSpan (e.g.
    // a header row: 1 cell spanning 2 + gridAfter 1) it inserts the filler
    // at the START, shoving the real cell right onto narrower columns (wrong
    // width + off-center). Emit explicit empty trailing cells so real cells
    // keep their left positions.
    const gridAfter = (row.gridAfter as number) ?? 0;
    if (gridAfter > 0) {
      const trailingType = (row.tableHeader as boolean) ? "tableHeader" : "tableCell";
      // gridAfter cells are empty trailing grid columns (no content). Give them
      // nil borders on every side so renderTableCellStyles emits border:none —
      // otherwise they pick up the Table-Grid default and draw a stray vertical
      // line at the row's right edge, showing up as an empty cell.
      const nilBorders = {
        top: { style: "nil" },
        right: { style: "nil" },
        bottom: { style: "nil" },
        left: { style: "nil" },
      };
      for (let c = 0; c < gridAfter; c++)
        cellNodes.push({
          type: trailingType,
          attrs: { borders: nilBorders },
          content: [{ type: "paragraph" }],
        });
      colIdx += gridAfter;
    }

    activeSpans = nextActiveSpans;

    const rowNode: JSONContent = { type: "tableRow" };
    if (Object.keys(rowAttrs).length > 0) rowNode.attrs = cleanAttrs(rowAttrs);
    if (cellNodes.length > 0) rowNode.content = cellNodes;

    content.push(rowNode);
  }

  const node: JSONContent = { type: "table" };
  if (Object.keys(attrs).length > 0) node.attrs = cleanAttrs(attrs);
  if (content.length > 0) node.content = content;

  return node;
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
  parseDocxBlock,
});
