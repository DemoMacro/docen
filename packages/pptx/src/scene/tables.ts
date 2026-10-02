// Graphic-frame table (p:graphicFrame a:tbl) projection: the merge/coverage
// grid walk, frame-border distribution, and the painter's normalized payload.

import { fillOpacityOf, measureEmu, outlineOf, solidFillOf } from "@docen/core/geometry";
import {
  emuToPx,
  stackBlocks,
  TextMeasurer,
  type LayoutBorderEdge,
  type LayoutDrawingMember,
} from "@docen/layout";
import type { CellBorderOptions, TableOptions, TableCellOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";
import { regionRulesAt, resolveTableStyle } from "./table-style";
import { textBlocks, type TextFieldContext } from "./text";

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
    if (!l) return { px: 0 };
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

/** One source cell at its grid slot (the addressing an editor needs to map
 *  a painted cell back to the object an edit writes into). */
export interface CellOrigin {
  cell: TableCellOptions;
  row: number;
  col: number;
  spanW: number;
  spanH: number;
}

/** The table's grid walk: source cells at their slots plus the column count
 *  the walk landed on. Every cell lands on the cursor. A "restart" merge
 *  field is the parser's spelling of the raw @hMerge/@vMerge="1" absorbed
 *  slot and consumes its cursor position; slots a rowSpan claims with no
 *  source cell (authoring omits the continuation) are skipped via coverage
 *  bookkeeping. An explicit "continue" (the raw ="0" not-merged cell) is a
 *  real cell. */
export function tableGridOf(table: TableOptions): { origins: CellOrigin[]; columns: number } {
  const nRows = table.rows.length;
  const covered: Set<number>[] = table.rows.map(() => new Set<number>());
  const origins: CellOrigin[] = [];
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
      for (let dr = 1; dr < spanH && r + dr < nRows; dr += 1)
        for (let dc = 0; dc < spanW; dc += 1) covered[r + dr]!.add(col + dc);
      origins.push({ cell, row: r, col, spanW, spanH });
      col += spanW;
      nCols = Math.max(nCols, col);
    }
  });
  return { origins, columns: nCols };
}

export function tableMember(
  table: TableOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
  context: TextFieldContext = {},
): LayoutDrawingMember {
  const { origins, columns: nCols } = tableGridOf(table);
  const style = resolveTableStyle(table, context.themeColors, context.tableStyles);

  // Declared column widths win; a table without them splits the frame evenly.
  const widths =
    table.columnWidths?.length === nCols
      ? table.columnWidths.map((w) => t.sx * emuOf(w))
      : Array.from({ length: nCols }, () => (t.sx * emuOf(table.width)) / Math.max(1, nCols));

  const frameBorders = table.borders;
  const nRows = table.rows.length;
  /** Style edges for one slot: rim edges from the region's own tokens, the
   *  inside tokens between cells. */
  const styleEdge = (
    rules: ReturnType<typeof regionRulesAt>,
    edge: "top" | "right" | "bottom" | "left",
    r: number,
    c: number,
    spanW: number,
    spanH: number,
  ) => {
    const b = rules.borders;
    if (!b) return undefined;
    if (edge === "top") return r === 0 ? b.top : b.insideH;
    if (edge === "bottom") return r + spanH >= nRows ? b.bottom : b.insideH;
    if (edge === "left") return c === 0 ? b.left : b.insideV;
    return c + spanW >= nCols ? b.right : b.insideV;
  };
  const declaredHeights = table.rows.map((row) => t.sy * emuOf(row.height));
  const measurer = context.metrics ? new TextMeasurer(context.metrics) : undefined;
  const rowNeed = table.rows.map(() => 0);
  if (measurer) {
    for (const origin of origins) {
      const rules = regionRulesAt(style, origin.row, origin.col, nRows, nCols);
      const innerWidth = Math.max(
        1,
        widths.slice(origin.col, origin.col + origin.spanW).reduce((sum, w) => sum + w, 0) -
          emuToPx(measureEmu(origin.cell.margins?.left) ?? CELL_INSET_EMU.left) -
          emuToPx(measureEmu(origin.cell.margins?.right) ?? CELL_INSET_EMU.right),
      );
      const blocks = textBlocks(
        { paragraphs: origin.cell.children, text: origin.cell.text },
        context,
      );
      for (const block of blocks) {
        if (block.kind !== "paragraph") continue;
        for (const inline of block.inline) {
          if (inline.kind !== "text") continue;
          if (rules.bold) inline.style.bold = true;
          if (rules.textColor) inline.style.color = rules.textColor;
        }
      }
      const need =
        t.sy *
        (stackBlocks(
          blocks,
          origin.cell.vertical === "vertical" || origin.cell.vertical === "vertical270"
            ? Math.max(
                1,
                declaredHeights
                  .slice(origin.row, Math.min(origin.row + origin.spanH, nRows))
                  .reduce((sum, height) => sum + height, 0) -
                  emuToPx(measureEmu(origin.cell.margins?.top) ?? CELL_INSET_EMU.top) -
                  emuToPx(measureEmu(origin.cell.margins?.bottom) ?? CELL_INSET_EMU.bottom),
              )
            : innerWidth,
          undefined,
          measurer,
        ).heightPx +
          emuToPx(measureEmu(origin.cell.margins?.top) ?? CELL_INSET_EMU.top) +
          emuToPx(measureEmu(origin.cell.margins?.bottom) ?? CELL_INSET_EMU.bottom));
      const last = Math.min(origin.row + origin.spanH - 1, nRows - 1);
      let declaredAbove = 0;
      for (let r = origin.row; r < last; r++) declaredAbove += Math.max(0, declaredHeights[r]!);
      rowNeed[last] = Math.max(rowNeed[last]!, need - declaredAbove);
    }
  }
  const rows = table.rows.map((row, r) => ({
    heightPx: Math.max(declaredHeights[r]!, rowNeed[r]!),
    cells: origins
      .filter((o) => o.row === r)
      .map(({ cell, col, spanW, spanH }) => {
        const rules = regionRulesAt(style, r, col, nRows, nCols);
        const fill = solidFillOf(cell.fill) ?? rules.fill;
        const opacity = fillOpacityOf(cell.fill);
        const edge = (key: "top" | "right" | "bottom" | "left") => {
          const direct = cell.borders?.[key];
          if (direct) return borderEdgeOf(direct);
          if (
            (key === "top" && r === 0) ||
            (key === "right" && col + spanW >= nCols) ||
            (key === "bottom" && r + spanH >= nRows) ||
            (key === "left" && col === 0)
          ) {
            const frame = frameBorders?.[key];
            if (frame) return borderEdgeOf(frame);
          }
          return styleEdge(rules, key, r, col, spanW, spanH);
        };
        const edges = {
          top: edge("top"),
          right: edge("right"),
          bottom: edge("bottom"),
          left: edge("left"),
        };
        const borders = Object.fromEntries(Object.entries(edges).filter(([, v]) => v));
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
          ...(cell.vertical
            ? {
                textVertical:
                  cell.vertical === "vertical270"
                    ? ("vertical270" as const)
                    : ("vertical" as const),
              }
            : {}),
          marginsPx: {
            left: m(cell.margins?.left, CELL_INSET_EMU.left),
            top: m(cell.margins?.top, CELL_INSET_EMU.top),
            right: m(cell.margins?.right, CELL_INSET_EMU.right),
            bottom: m(cell.margins?.bottom, CELL_INSET_EMU.bottom),
          },
          blocks: (() => {
            const blocks = textBlocks({ paragraphs: cell.children, text: cell.text }, context);
            if (!rules.bold && !rules.textColor) return blocks;
            for (const block of blocks) {
              if (block.kind !== "paragraph") continue;
              for (const inline of block.inline) {
                if (inline.kind !== "text") continue;
                if (rules.bold) inline.style.bold = true;
                if (rules.textColor) inline.style.color = rules.textColor;
              }
            }
            return blocks;
          })(),
        };
      }),
  }));

  return {
    kind: "table",
    x: t.sx * emuOf(table.x) + t.dx,
    y: t.sy * emuOf(table.y) + t.dy,
    width: widths.reduce((a, w) => a + w, 0),
    height: rows.reduce((sum, row) => sum + row.heightPx, 0),
    ...(childPath ? { childPath } : {}),
    table: { columnWidthsPx: widths, rows },
  };
}
