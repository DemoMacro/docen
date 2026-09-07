/**
 * Table interaction geometry — the math behind canvas table resizing,
 * border hit detection, and quick-insert controls. Pure functions on plain
 * coordinates and width/height arrays.
 */

import type { TableZone } from "./caret-map";

/** 1 px = 15 twips (96 DPI standard: 1440 twips / 96 px = 15 twips/px). */
export const PX_TO_TWIPS = 15;
export const TWIPS_TO_PX = 1 / 15;

/** Minimum column width in twips: 0.5 inch (720 twips = 48 px). */
export const MIN_COL_TWIPS = 720;

/** Minimum row height in twips: 9 pt (180 twips = 12 px). */
export const MIN_ROW_TWIPS = 180;

/** Hit-test margin for table borders in page px. */
export const BORDER_GRAB_PX = 4;

/** Hit-test radius for the bottom-right corner resize handle in page px. */
export const CORNER_HANDLE_RADIUS_PX = 7;

export type TableHit =
  | { kind: "col-border"; colIndex: number }
  | { kind: "row-border"; rowIndex: number }
  | { kind: "corner-handle" }
  | { kind: "quick-col"; colIndex: number }
  | { kind: "quick-row"; rowIndex: number };

/**
 * Hit-test a table zone against a page-local coordinate (lx, ly).
 * Returns the hit kind (col-border, row-border, corner-handle, or quick-insert)
 * or null if the pointer is not over any interactive handle.
 */
export function hitTableElements(zone: TableZone, lx: number, ly: number): TableHit | null {
  const tx = lx - zone.xPx;
  const ty = ly - zone.yPx;

  // 1. Bottom-right corner resize handle
  const distCornerX = Math.abs(tx - zone.widthPx);
  const distCornerY = Math.abs(ty - zone.heightPx);
  if (distCornerX <= CORNER_HANDLE_RADIUS_PX && distCornerY <= CORNER_HANDLE_RADIUS_PX) {
    return { kind: "corner-handle" };
  }

  // 2. Quick-insert column (+) button area: hovering just above the table top edge
  if (ty >= -20 && ty <= 2) {
    for (let i = 1; i < zone.colEdges.length; i += 1) {
      if (Math.abs(tx - zone.colEdges[i]!) <= 10) {
        return { kind: "quick-col", colIndex: i };
      }
    }
  }

  // 3. Quick-insert row (+) button area: hovering just to the left of the table
  if (tx >= -20 && tx <= 2) {
    for (let r = 1; r < zone.rowEdges.length; r += 1) {
      if (Math.abs(ty - zone.rowEdges[r]!) <= 10) {
        return { kind: "quick-row", rowIndex: r };
      }
    }
  }

  // 4. Column borders (vertical lines): index 1 through colEdges.length - 1
  // (Interior borders + right border). Edge index `i` sits between col `i - 1` and col `i`.
  if (ty >= -BORDER_GRAB_PX && ty <= zone.heightPx + BORDER_GRAB_PX) {
    for (let i = 1; i < zone.colEdges.length; i += 1) {
      if (Math.abs(tx - zone.colEdges[i]!) <= BORDER_GRAB_PX) {
        return { kind: "col-border", colIndex: i };
      }
    }
  }

  // 5. Row borders (horizontal lines): index 1 through rowEdges.length - 1
  // (Bottom border of row `r - 1`).
  if (tx >= -BORDER_GRAB_PX && tx <= zone.widthPx + BORDER_GRAB_PX) {
    for (let r = 1; r < zone.rowEdges.length; r += 1) {
      if (Math.abs(ty - zone.rowEdges[r]!) <= BORDER_GRAB_PX) {
        return { kind: "row-border", rowIndex: r };
      }
    }
  }

  return null;
}

/**
 * Calculates new column widths (in twips) when dragging a column border.
 *
 * @param initialWidths Original column widths array (twips).
 * @param colIndex 1-based border index. 1 = between col 0 and 1; N = right border.
 * @param dxTwips Mouse movement delta in twips.
 * @param options Shift key and minimum width configuration.
 */
export function calcColumnResize(
  initialWidths: readonly number[],
  colIndex: number,
  dxTwips: number,
  options: { shiftKey?: boolean; minWidthTwips?: number } = {},
): number[] {
  const minWidth = options.minWidthTwips ?? MIN_COL_TWIPS;
  const nCols = initialWidths.length;
  if (nCols === 0 || colIndex < 1 || colIndex > nCols) {
    return [...initialWidths];
  }

  const widths = [...initialWidths];

  // Case A: Rightmost border (changes table total width)
  if (colIndex === nCols) {
    const lastIdx = nCols - 1;
    widths[lastIdx] = Math.max(minWidth, (widths[lastIdx] ?? 0) + dxTwips);
    return widths;
  }

  // Case B: Shift pressed on interior border — shifts all subsequent columns rightward
  if (options.shiftKey) {
    const leftIdx = colIndex - 1;
    widths[leftIdx] = Math.max(minWidth, (widths[leftIdx] ?? 0) + dxTwips);
    return widths;
  }

  // Case C: Standard Word drag on interior border — redistributes width between col i-1 and col i
  const leftIdx = colIndex - 1;
  const rightIdx = colIndex;
  const wLeft = widths[leftIdx] ?? 0;
  const wRight = widths[rightIdx] ?? 0;
  const sum = wLeft + wRight;

  const targetLeft = wLeft + dxTwips;
  const clampedLeft = Math.max(minWidth, Math.min(sum - minWidth, targetLeft));
  const clampedRight = sum - clampedLeft;

  widths[leftIdx] = clampedLeft;
  widths[rightIdx] = clampedRight;
  return widths;
}

/**
 * Calculates new row height when dragging a row's bottom border.
 *
 * @param initialHeightTwips Initial height of the row in twips (or natural height).
 * @param dyTwips Vertical drag delta in twips.
 * @param options Modifier flags (e.g. Alt key for "exact" rule) and minimum height.
 */
export function calcRowResize(
  initialHeightTwips: number,
  dyTwips: number,
  options: { isAlt?: boolean; minHeightTwips?: number } = {},
): { rule: "atLeast" | "exact"; value: number } {
  const minHeight = options.minHeightTwips ?? MIN_ROW_TWIPS;
  const target = Math.max(minHeight, initialHeightTwips + dyTwips);
  return {
    rule: options.isAlt ? "exact" : "atLeast",
    value: Math.round(target),
  };
}

/**
 * Scales all column widths proportionally when dragging the bottom-right corner handle.
 *
 * @param initialWidths Original column widths array (twips).
 * @param scaleX Horizontal scaling factor (> 0).
 * @param minWidthTwips Minimum column width.
 */
export function calcTableScale(
  initialWidths: readonly number[],
  scaleX: number,
  minWidthTwips = MIN_COL_TWIPS,
): number[] {
  if (initialWidths.length === 0 || !Number.isFinite(scaleX) || scaleX <= 0) {
    return [...initialWidths];
  }
  return initialWidths.map((w) => Math.max(minWidthTwips, Math.round(w * scaleX)));
}
