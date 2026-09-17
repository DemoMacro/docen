// Graphic-frame table (p:graphicFrame a:tbl) projection: the merge/coverage
// grid walk, frame-border distribution, and the painter's normalized payload.

import { fillOpacityOf, measureEmu, outlineOf, solidFillOf } from "@docen/core/geometry";
import { emuToPx, type LayoutBorderEdge, type LayoutDrawingMember } from "@docen/layout";
import type { CellBorderOptions, TableOptions, TableCellOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";
import { textBlocks } from "./text";

// DrawingML table cell default margins (marL/marR 0.1", marT/marB 0.05" —
// the same insets a bodyPr carries).
const CELL_INSET_EMU = { left: 91440, top: 45720, right: 91440, bottom: 45720 };

/** prstDash tokens → the border-style tokens the painter's dash map knows;
 *  unmapped composites (lgDash, wide composites) render solid. */
const BORDER_DASH: Record<string, string> = {
  dash: "dashed",
  sysDash: "dashed",
  lgDash: "dashed",
  dot: "dotted",
  sysDot: "dotted",
  dashSmallGap: "dashSmallGap",
  dashDot: "dashDot",
  lgDashDot: "dashDot",
  lgDashDotDot: "dashDotDot",
};

/** One cell border → the edge record the painter strokes. A noFill line (an
 *  explicitly erased edge) maps to nothing. */
function borderEdgeOf(b: CellBorderOptions): LayoutBorderEdge | undefined {
  if (b.outline) {
    const l = outlineOf(b.outline);
    if (!l) return undefined;
    return {
      px: l.px,
      ...(l.color ? { color: l.color } : {}),
      ...(l.dash && BORDER_DASH[l.dash] ? { style: BORDER_DASH[l.dash] } : {}),
    };
  }
  // Sugar form: 1 pt when the width is unspecified (PowerPoint's table
  // border default).
  return {
    px: emuToPx(measureEmu(b.width) ?? 12700),
    ...(solidFillOf(b.color) ? { color: solidFillOf(b.color) } : {}),
    ...(b.dashStyle && BORDER_DASH[b.dashStyle] ? { style: BORDER_DASH[b.dashStyle] } : {}),
  };
}

export function tableMember(
  table: TableOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember {
  // Resolve the grid: every cell lands on the cursor. A "restart" merge field
  // is the parser's spelling of the raw @hMerge/@vMerge="1" absorbed slot and
  // consumes its cursor position; slots a rowSpan claims with no source cell
  // (authoring omits the continuation) are skipped via coverage bookkeeping.
  // An explicit "continue" (the raw ="0" not-merged cell) is a real cell.
  const nRows = table.rows.length;
  const covered: Set<number>[] = table.rows.map(() => new Set<number>());
  const origins: {
    cell: TableCellOptions;
    row: number;
    col: number;
    spanW: number;
    spanH: number;
  }[] = [];
  let nCols = 0;
  table.rows.forEach((row, r) => {
    let col = 0;
    for (const cell of row.cells) {
      if (cell.horizontalMerge === "restart" || cell.verticalMerge === "restart") {
        col += 1;
        nCols = Math.max(nCols, col);
        continue;
      }
      while (covered[r]!.has(col)) col += 1;
      const spanW = cell.columnSpan ?? 1;
      const spanH = cell.rowSpan ?? 1;
      for (let k = 1; k < spanH && r + k < nRows; k++) covered[r + k]!.add(col);
      origins.push({ cell, row: r, col, spanW, spanH });
      col += spanW;
      nCols = Math.max(nCols, col);
    }
  });

  // Declared column widths win; a table without them splits the frame evenly.
  const widths = table.columnWidths?.length
    ? table.columnWidths.map((w) => t.sx * emuOf(w))
    : Array.from({ length: nCols }, () => (t.sx * emuOf(table.width)) / Math.max(1, nCols));

  const frameBorders = table.borders;
  const rows = table.rows.map((row, r) => ({
    heightPx: t.sy * emuOf(row.height),
    cells: origins
      .filter((o) => o.row === r)
      .map(({ cell, col, spanW, spanH }) => {
        const fill = solidFillOf(cell.fill);
        const opacity = fillOpacityOf(cell.fill);
        const borders = {
          ...(cell.borders?.top ? { top: borderEdgeOf(cell.borders.top) } : {}),
          ...(cell.borders?.right ? { right: borderEdgeOf(cell.borders.right) } : {}),
          ...(cell.borders?.bottom ? { bottom: borderEdgeOf(cell.borders.bottom) } : {}),
          ...(cell.borders?.left ? { left: borderEdgeOf(cell.borders.left) } : {}),
        };
        // Frame-level borders spread onto the rim cells (the stringify side's
        // distributeBorders contract) where the cell declares none of its own.
        if (frameBorders) {
          if (r === 0 && frameBorders.top && !borders.top)
            borders.top = borderEdgeOf(frameBorders.top);
          if (col + spanW >= nCols && frameBorders.right && !borders.right)
            borders.right = borderEdgeOf(frameBorders.right);
          if (r + spanH >= nRows && frameBorders.bottom && !borders.bottom)
            borders.bottom = borderEdgeOf(frameBorders.bottom);
          if (col === 0 && frameBorders.left && !borders.left)
            borders.left = borderEdgeOf(frameBorders.left);
        }
        const m = (v: unknown, def: number) => emuToPx(measureEmu(v) ?? def);
        return {
          col,
          row: r,
          spanW,
          spanH,
          ...(fill ? { fill } : {}),
          ...(opacity != null ? { opacity } : {}),
          ...(Object.keys(borders).length > 0 ? { borders } : {}),
          ...(cell.verticalAlign === "center" || cell.verticalAlign === "bottom"
            ? { anchor: cell.verticalAlign }
            : {}),
          marginsPx: {
            left: m(cell.margins?.left, CELL_INSET_EMU.left),
            top: m(cell.margins?.top, CELL_INSET_EMU.top),
            right: m(cell.margins?.right, CELL_INSET_EMU.right),
            bottom: m(cell.margins?.bottom, CELL_INSET_EMU.bottom),
          },
          blocks: textBlocks({ paragraphs: cell.children, text: cell.text }),
        };
      }),
  }));

  return {
    kind: "table",
    x: t.sx * emuOf(table.x) + t.dx,
    y: t.sy * emuOf(table.y) + t.dy,
    width: widths.reduce((a, w) => a + w, 0),
    height: t.sy * emuOf(table.height),
    ...(childPath ? { childPath } : {}),
    table: { columnWidthsPx: widths, rows },
  };
}
