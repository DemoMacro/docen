// Table projection: direct tblPr over the table style — widths, insets,
// borders, per-row height/header/cantSplit, and vertical merges expanded
// into the engine's rowspan shape.

import {
  twipToPx,
  type LayoutBlock,
  type LayoutBorderEdge,
  type LayoutCellInsets,
  type LayoutTable,
  type LayoutTableWidth,
} from "@docen/layout";
import type { TableCellOptions, TableOptions } from "@office-open/docx";

import { indexTableStyles } from "../../style-cascade";
import type { ProjectContext } from "./context";
import { eighthPtToPx, isRecord, measureTwip, num, type LayoutCell, type Rec } from "./guards";
import { projectChild } from "./page";

// ── table projection ──

function toTableWidth(w: unknown): LayoutTableWidth | undefined {
  if (!isRecord(w)) return undefined;
  if (w.type === "auto" || w.type === "nil") return undefined;
  if (w.type === "percent") {
    const size = w.size;
    const pct =
      typeof size === "string" && size.endsWith("%") ? Number(size.slice(0, -1)) : num(size);
    if (pct != null && Number.isFinite(pct)) return { type: "percent", percent: pct };
    return undefined;
  }
  const tw = measureTwip(w.size);
  return tw != null && tw > 0 ? { type: "px", px: twipToPx(tw) } : undefined;
}

function toCellInsets(m: unknown): LayoutCellInsets | undefined {
  if (!isRecord(m)) return undefined;
  const side = (v: unknown): number | undefined => {
    const size = isRecord(v) ? measureTwip(v.size) : undefined;
    return size != null ? twipToPx(size) : undefined;
  };
  const insets = {
    top: side(m.top),
    right: side(m.right),
    bottom: side(m.bottom),
    left: side(m.left),
  };
  return insets.top != null || insets.right != null || insets.bottom != null || insets.left != null
    ? insets
    : undefined;
}

/** Word's application default when neither the table nor its style declares
 *  w:tblCellMar: 108 twips left/right, 0 top/bottom. Without it cells wrap at
 *  the full column width and their text paints over the borders. */
const WORD_DEFAULT_CELL_INSETS: LayoutCellInsets = {
  left: twipToPx(108),
  right: twipToPx(108),
};

type CellBorders = NonNullable<LayoutCell["borders"]>;

/** One w:tcBorders/w:tblBorders edge → px + color (nil/none survive as declared
 *  zero-weight edges; the conflict resolver skips them). */
function toBorderEdge(v: unknown): LayoutBorderEdge | undefined {
  if (!isRecord(v)) return undefined;
  const size = num(v.size);
  const color = typeof v.color === "string" && v.color !== "auto" ? v.color : undefined;
  return {
    style: typeof v.style === "string" ? v.style : undefined,
    px: size != null ? eighthPtToPx(size) : undefined,
    color,
  };
}

function toBorders(b: unknown): CellBorders | undefined {
  if (!isRecord(b)) return undefined;
  const out = {
    top: toBorderEdge(b.top),
    right: toBorderEdge(b.right),
    bottom: toBorderEdge(b.bottom),
    left: toBorderEdge(b.left),
  };
  return out.top || out.right || out.bottom || out.left ? out : undefined;
}

/** w:tblBorders → the engine's table-level defaults, merging the direct
 *  tblPr borders over the table style's per side. */
function toTableBorders(direct: unknown, styleTable: unknown): LayoutTable["borders"] | undefined {
  const d = isRecord(direct) ? direct : undefined;
  const s = isRecord(styleTable)
    ? isRecord(styleTable.borders)
      ? styleTable.borders
      : undefined
    : undefined;
  if (!d && !s) return undefined;
  const edge = (side: string): LayoutBorderEdge | undefined =>
    toBorderEdge(d?.[side]) ?? toBorderEdge(s?.[side]);
  const out = {
    top: edge("top"),
    bottom: edge("bottom"),
    left: edge("left"),
    right: edge("right"),
    insideHorizontal: edge("insideHorizontal"),
    insideVertical: edge("insideVertical"),
  };
  return out.top ||
    out.bottom ||
    out.left ||
    out.right ||
    out.insideHorizontal ||
    out.insideVertical
    ? out
    : undefined;
}

function projectCell(c: TableCellOptions, ctx: ProjectContext, rowspan?: number): LayoutCell {
  const shd = isRecord(c.shading) ? c.shading : undefined;
  const fill =
    shd && typeof shd.fill === "string" && shd.fill !== "auto" && shd.type !== "nil"
      ? shd.fill
      : undefined;
  return {
    colspan: c.columnSpan,
    rowspan: rowspan ?? 1,
    insets: toCellInsets(c.margins),
    borders: toBorders(c.borders),
    fill,
    verticalAlign:
      c.verticalAlign === "center" || c.verticalAlign === "bottom" ? c.verticalAlign : undefined,
    blocks: c.children
      .map((child) => projectChild(child, ctx))
      .filter((b): b is LayoutBlock => b !== null),
  };
}

/** Expand OOXML vertical merges into the layout's rowspan shape — the single
 *  projection point where the two models meet. A `restart` cell absorbs every
 *  `continue` cell below it in the same grid columns: the continuation rows
 *  drop those cells (OOXML gives them just an empty <w:p>/) and the restart's
 *  rowspan counts them. Returns the merged-cell rowspan per restart cell. */
function collectRowSpans(rows: { cells: unknown[] }[]): Map<TableCellOptions, number> {
  const spans = new Map<TableCellOptions, number>();
  // Grid column → the restart cell currently absorbing continuations below.
  const open = new Map<number, TableCellOptions>();
  for (const row of rows) {
    let col = 0;
    for (const raw of row.cells) {
      if (!isRecord(raw) || !("children" in raw)) continue;
      const cell = raw as unknown as TableCellOptions;
      const span = cell.columnSpan ?? 1;
      if (cell.verticalMerge === "continue") {
        const owner = open.get(col);
        if (owner) spans.set(owner, (spans.get(owner) ?? 1) + 1);
      } else if (cell.verticalMerge === "restart") {
        spans.set(cell, 1);
        for (let c = col; c < col + span; c++) open.set(c, cell);
        col += span;
        continue;
      } else {
        for (let c = col; c < col + span; c++) open.delete(c);
        col += span;
        continue;
      }
      col += span;
    }
  }
  return spans;
}

export function projectTable(t: TableOptions, ctx: ProjectContext): LayoutTable {
  // Only cell rows project; vMerge continuation cells fold into their restart.
  const cellRows = (t.rows ?? []).filter(
    (row): row is Extract<(typeof t.rows)[number], { cells: unknown[] }> =>
      "cells" in row && Array.isArray(row.cells),
  );
  const rowSpans = collectRowSpans(cellRows);
  const rows: LayoutTable["rows"] = [];
  for (const row of cellRows) {
    const trHeight: Rec = isRecord(row.height) ? row.height : {};
    const heightValue = measureTwip(trHeight.value);
    const height =
      heightValue != null && heightValue > 0
        ? {
            rule: trHeight.rule === "exact" ? ("exact" as const) : ("atLeast" as const),
            px: twipToPx(heightValue),
          }
        : undefined;
    rows.push({
      cells: row.cells
        .filter((cell): cell is TableCellOptions => "children" in cell)
        .filter((cell) => cell.verticalMerge !== "continue")
        .map((cell) => projectCell(cell, ctx, rowSpans.get(cell))),
      height,
      tableHeader: row.tableHeader || undefined,
      cantSplit: row.cantSplit || undefined,
    });
  }

  const columnWidthsPx = t.columnWidths?.map((w) => twipToPx(measureTwip(w) ?? 0));
  const styleTable = t.style ? indexTableStyles(ctx.styles).get(t.style)?.table : undefined;
  const alignment = t.alignment ?? styleTable?.alignment;
  return {
    kind: "table",
    width: toTableWidth(t.width),
    layout: t.layout,
    align:
      alignment === "center"
        ? "center"
        : alignment === "right" || alignment === "end"
          ? "right"
          : undefined,
    columnWidthsPx: columnWidthsPx && columnWidthsPx.length > 0 ? columnWidthsPx : undefined,
    cellInsets: toCellInsets(t.margins) ?? WORD_DEFAULT_CELL_INSETS,
    borders: toTableBorders(t.borders, styleTable),
    rows,
  };
}
