import type { LayoutBlock, LayoutBorderEdge } from "./block";

/** Cell insets in px: per-cell w:tcMar, or the table's w:tblCellMar default
 *  (cells inherit per side — the adapter or the engine resolves cell ?? table). */
export interface LayoutCellInsets {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/** CT_TblBorders: the table-level edges a cell's own w:tcBorders side falls
 *  back to — outer edges for the grid's rim, inside edges for shared lines. */
export interface LayoutTableBorders {
  top?: LayoutBorderEdge;
  bottom?: LayoutBorderEdge;
  left?: LayoutBorderEdge;
  right?: LayoutBorderEdge;
  insideHorizontal?: LayoutBorderEdge;
  insideVertical?: LayoutBorderEdge;
}

export interface LayoutTableCell {
  colspan?: number;
  rowspan?: number;
  /** Per-spanned-column widths in px (w:tcW resolved); absent → grid share. */
  widthPx?: number;
  insets?: LayoutCellInsets;
  borders?: Partial<Record<"top" | "right" | "bottom" | "left", LayoutBorderEdge>>;
  /** Cell shading (w:shd @w:fill), hex RRGGBB. */
  fill?: string;
  /** w:vAlign — the content's placement when the row is taller than it. */
  verticalAlign?: "top" | "center" | "bottom";
  blocks: LayoutBlock[];
}

export interface LayoutTableRow {
  cells: LayoutTableCell[];
  /** w:trHeight resolved: atLeast floors the row, exact fixes it (content
   *  overflows but the row does not grow). */
  height?: { rule: "atLeast" | "exact"; px: number };
  /** w:tblHeader — leading rows repeat on every page the table splits onto
   *  (only a contiguous prefix from the first row counts). */
  tableHeader?: boolean;
  /** w:cantSplit — the row moves whole to the next page instead of splitting
   *  mid-content (a row taller than a page still force-splits — Word clips
   *  nothing). */
  cantSplit?: boolean;
}

export type LayoutTableWidth =
  | { type: "percent"; percent: number } // 0-100
  | { type: "px"; px: number }; // w:tblW dxa resolved

export interface LayoutTable {
  kind: "table";
  /** Absent width = auto: fill the containing flow width. */
  width?: LayoutTableWidth;
  /** w:tblPr/w:jc — the table box's placement inside the flow column. A table
   *  wider than the column centers into the margins (negative offset). */
  align?: "left" | "center" | "right";
  /** tblGrid column widths in px, scaled proportionally to the effective
   *  table width (Word scales the grid to tblW, never to the raw sum). */
  columnWidthsPx?: number[];
  /** Table-level default insets (w:tblCellMar) a cell without its own w:tcMar
   *  inherits, per side. */
  cellInsets?: LayoutCellInsets;
  /** Table-level border defaults (w:tblBorders, style chain resolved):
   *  the renderer falls a cell's missing edge back to these per side. */
  borders?: LayoutTableBorders;
  rows: LayoutTableRow[];
}
