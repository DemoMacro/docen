import type {
  BorderOptions,
  SectionChild,
  TableOptions,
  TablePropertiesOptions,
  TableWidthProperties,
} from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { Node } from "@tiptap/core";

import { allBordersNone, cleanAttrs } from "../converters/styles";
import { mergeTableStyleProps } from "../style-cascade";
import type { ParseBlockRule, ResolveContext } from "./types";
import {
  attrNative,
  alignmentFromElement,
  alignmentToCss,
  bordersFromElement,
  type DocxAttrSpec,
  shadingFromElement,
  shadingToCss,
  tableFloatToCss,
  twipToCss,
  renderBorderCSS,
} from "./utils";

/**
 * Table node with nested office-open attrs — fully custom (no upstream table
 * extension).
 *
 * Attrs mirror TableOptions (width/float/layout/borders/alignment/margins/indent/
 * cellSpacing/tableLook/columnWidths/etc.). DOCX round-trip is near-identity:
 * renderDocx/parseDocx pass attrs through (omitting only `rows`, which DocxManager
 * rebuilds from the row/cell nodes). CSS conversion happens only in renderHTML.
 */

// ── DOCX serialization (near-identity: attrs mirror TableOptions minus rows) ──

/** Structural keys not mirrored as attrs: `rows` is rebuilt by DocxManager
 *  (compileTableNode walks the row nodes); `columnWidthsRevision` (a tblGrid
 *  change revision) is skipped at parse too — the editor's tblGrid edits are
 *  new revisions, not replays of the old one. Keep in sync with SKIP_KEYS
 *  below (SKIP = never enters opts; these = never enter attrs). */
const SKIP_KEYS = new Set(["rows", "columnWidthsRevision"]);

/** The attr key set the table node declares — every TablePropertiesOptions key
 *  (TableOptions' own properties incl. revision) plus the tblGrid widths.
 *  Excluded: `rows` (rebuilt by DocxManager), `columnWidthsRevision` (skipped
 *  at parse too), and the core base's 6 band flags — docx models those in
 *  `tableLook` (w:tblLook); they are stringify-only authoring shorthand and
 *  parse never emits them as top-level keys (descriptor.ts tableLookForEmit). */
type TableAttrKey = Exclude<
  keyof TablePropertiesOptions | keyof TableOptions,
  | "rows"
  | "columnWidthsRevision"
  | "firstRow"
  | "lastRow"
  | "firstCol"
  | "lastCol"
  | "bandRow"
  | "bandCol"
>;

/** office-open table attr mirror, satisfies-guarded against keyof drift (same
 *  contract as docxParagraphAttrs in utils.ts). */
const docxTableAttrs = {
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
  // Table-level property revision (w:tblPrChange).
  revision: attrNative(),
} satisfies Record<TableAttrKey, DocxAttrSpec>;

export function renderDocx(node: JSONContent): Partial<TableOptions> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (SKIP_KEYS.has(key)) continue;
    if (value !== null && value !== undefined) opts[key] = value;
  }
  return opts;
}

export function parseDocx(opts: TableOptions): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(opts)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Block parse rule (resolve: SectionChild → table node) ──

/**
 * Declarative block parse rule: recognize a table SectionChild and resolve it
 * as a Tiptap table node (cell attrs pass through verbatim — columnSpan/
 * verticalMerge/width included; the table style's tblBorders/tblCellMar merged
 * in, insideH/V grid lines pushed onto cells, gridAfter as trailing
 * nil-bordered cells). DocxManager dispatches every SectionChild through this
 * rule before the paragraph/passthrough fallbacks. */
export const parseDocxBlock: ParseBlockRule<Extract<SectionChild, { table: TableOptions }>> = {
  match: (child): child is Extract<SectionChild, { table: TableOptions }> => "table" in child,
  convert: (child, ctx) => resolveTable(child.table, ctx),
};

/** Resolve a table SectionChild into a Tiptap table node. A cell is itself a
 *  SectionChild[] block stream, resolved recursively via ctx. */
function resolveTable(tableOpts: TableOptions, ctx: ResolveContext): JSONContent {
  const attrs = ctx.parseNodeAttrs("table", tableOpts);
  const content: JSONContent[] = [];

  // The walk below reads rows/cells reflectively (attrs parse + span
  // bookkeeping) and treats the sdt/customXml/marker row variants the same as
  // cell rows, so it views them as attribute records.
  const rows = tableOpts.rows as unknown as Record<string, unknown>[];

  // Pull the referenced table style's tblBorders/tblCellMar in: office-open
  // leaves table.borders/margins reflecting only the table's own tblPr, so
  // a "Table Grid" table (borders defined in the style) would render no grid
  // without this. The table's own real borders win; the style fills the gap
  // when the table's are all none/nil.
  const styleProps = mergeTableStyleProps(ctx.styles?.tableStyles, tableOpts.style ?? null);
  if (styleProps.borders && allBordersNone(tableOpts.borders)) {
    tableOpts = { ...tableOpts, borders: styleProps.borders };
  }
  if (styleProps.margins && tableOpts.margins == null) {
    tableOpts = { ...tableOpts, margins: styleProps.margins };
  }

  // Table-level default cell insets (w:tblCellMar). office-open exposes them
  // on both `cellMargin` and `margins`; a cell inherits them unless it carries
  // its own tcMar. Push the default onto cells without tcMar so render
  // (renderTableCellStyles) and the paginator (cellVerticalOverhead) read ONE
  // effective source (cell.attrs.margins) instead of each falling back to the
  // table. compileTableCellNode drops a cell tcMar equal to this default to
  // keep the regenerated docx in its table-level form (near-identity round-trip).
  const tableCellMargins = tableOpts.margins ?? null;
  // Table-level inside grid lines (tblBorders.insideHorizontal/insideVertical).
  // In CSS border-collapse the interior grid belongs to cells, not the <table>
  // element, so a REAL insideH/V is pushed onto cell sides lacking their own
  // tcBorder (below). none/nil is skipped — a table that merely LACKS inner
  // grid lines must not have `border:none` stamped on every cell, or it would
  // suppress the editor's Table-Grid fallback default for borderless tables.
  // Edge cells' outer sides overlap the table's own border under
  // border-collapse (thicker wins), matching OOXML outer-vs-inner semantics.
  const tableBorders = tableOpts.borders ?? null;
  const realBorder = (bd: unknown): BorderOptions | null => {
    const b = bd as BorderOptions | null | undefined;
    return b && b.style && b.style !== "none" && b.style !== "nil" ? b : null;
  };
  const insideH = realBorder(tableBorders?.insideHorizontal);
  const insideV = realBorder(tableBorders?.insideVertical);

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
    const cells = (row.cells ?? []) as unknown as Record<string, unknown>[];
    const cellNodes: JSONContent[] = [];

    // A `continue` cell (vMerge) is a real node — attrs pass through verbatim
    // (verticalMerge: "continue"); the layout engine folds it into the restart
    // cell's rowspan at the single projection point.
    for (const cell of cells) {
      const cellAttrs = ctx.parseNodeAttrs("tableCell", cell as unknown as Record<string, unknown>);

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

      const cellNode: JSONContent = { type: "tableCell" };
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

      cellNodes.push(cellNode);
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
          type: "tableCell",
          attrs: { borders: nilBorders },
          content: [{ type: "paragraph" }],
        });
    }

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

export const Table = Node.create({
  name: "table",
  group: "block",
  content: "tableRow+",

  addAttributes() {
    return docxTableAttrs;
  },

  parseHTML() {
    return [{ tag: "table" }];
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

    // A floating table (w:tblpPr) renders as CSS float + margins. When float is
    // active it owns the table's margins and width, so the alignment margins and
    // the default 100% width stand down below. When float degrades to [] (page/
    // margin anchor, or center/inside/outside), the table renders as a normal
    // block and alignment/width apply as before.
    const floatStyles = tableFloatToCss(a.float);

    if (floatStyles.length === 0) {
      const align = alignmentToCss(a.alignment as string | null | undefined);
      if (align === "center") {
        styles.push("margin-left:auto", "margin-right:auto");
      } else if (align === "right") {
        styles.push("margin-left:auto", "margin-right:0");
      } else if (align) {
        styles.push("margin-left:0", "margin-right:auto");
      }
    }

    if (a.layout === "fixed") styles.push("table-layout:fixed");

    // Table width: pct → percentage; dxa/numeric → twips; "auto" or no
    // <w:tblW> → fill the page content area. Word's "auto" tables size to the
    // text column (full-width unless content is sparse); with no width the
    // browser would shrink the table to its content, so match Word with 100%.
    if (a.width && typeof a.width === "object") {
      const w = a.width as TableWidthProperties;
      const numSize = typeof w.size === "string" ? parseFloat(w.size) : w.size;
      if (w.type === "percent") {
        if (typeof w.size === "string" && w.size.includes("%")) {
          // office-open keeps the percentage literal verbatim ("100%" = 100%) —
          // it is NOT fiftieths-of-a-percent, so do not divide by 50.
          styles.push(`width:${w.size}`);
        } else if (!Number.isNaN(numSize)) {
          // office-open normalizes pct to a percentage number (0-100) via
          // widthFiftiethsToPct (fiftieths / 50) on parse — 99.96 ≈ 100%. Use
          // it directly; dividing by 50 again collapses the table to a sliver.
          styles.push(`width:${numSize}%`);
        }
      } else if (w.type === "auto") {
        styles.push(floatStyles.length ? "width:auto" : "width:100%");
      } else if (numSize != null) {
        const css = twipToCss(numSize);
        if (css) styles.push(`width:${css}`);
      }
    } else {
      styles.push(floatStyles.length ? "width:auto" : "width:100%");
    }
    // Cap dxa tables at the page text column. A tblW in twips can exceed the
    // content area (Word's table-width vs grid-column-sum mismatch — e.g. a
    // stray narrow trailing grid column inflates tblW past the text column),
    // which overflows the fixed page box on the right. max-width:100% shrinks
    // only over-wide tables; pct/auto/100% widths are already ≤100% so this is
    // a no-op for them.
    styles.push("max-width:100%");

    if (a.indent && typeof a.indent === "object") {
      const ind = a.indent as TableWidthProperties;
      if (typeof ind.size === "number") {
        const css = twipToCss(ind.size);
        if (css) styles.push(`margin-left:${css}`);
      }
    }

    if (a.cellSpacing && typeof a.cellSpacing === "object") {
      // TableWidthProperties: { size, type } — office-open's cellSpacingStr reads
      // opts.size (comments-BKbOd_pv.mjs:7451). The previous `cs.value` read never
      // matched, so <w:tblCellSpacing> was silently dropped from the editor render.
      const cs = a.cellSpacing as { size?: number; type?: string };
      if (cs.size != null) {
        const css = twipToCss(cs.size);
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

    // colgroup columns from the tblGrid (columnWidths) — the table's real
    // column structure.
    const gridPx = ((a.columnWidths as Array<number> | null) ?? []).map((w) =>
      Math.round((w || 0) / 15),
    );

    // Float styles last so they win any same-property duel with the alignment
    // margins above (which only run on the floatStyles.length === 0 path).
    for (const s of floatStyles) styles.push(s);

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
