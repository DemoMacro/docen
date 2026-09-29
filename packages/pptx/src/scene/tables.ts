// Graphic-frame table (p:graphicFrame a:tbl) projection: the merge/coverage
// grid walk, frame-border distribution, and the painter's normalized payload.

import { fillOpacityOf, measureEmu, outlineOf, solidFillOf } from "@docen/core/geometry";
import { emuToPx, type LayoutBorderEdge, type LayoutDrawingMember } from "@docen/layout";
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
      for (let k = 1; k < spanH && r + k < nRows; k++) covered[r + k]!.add(col);
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
  const widths = table.columnWidths?.length
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
  // OOXML permits h="0" for auto-sized rows. PowerPoint still lays the frame
  // out over its graphicFrame extent; keep that geometry for painting and the
  // editor's cell hit-testing instead of collapsing every row onto y=0.
  const frameHeightPx = t.sy * emuOf(table.height);
  const declaredHeights = table.rows.map((row) => t.sy * emuOf(row.height));
  const autoRows = declaredHeights.filter((height) => height <= 0).length;
  const fixedHeightPx = declaredHeights.reduce((sum, height) => sum + Math.max(0, height), 0);
  const autoHeightPx =
    frameHeightPx > fixedHeightPx && autoRows > 0 ? (frameHeightPx - fixedHeightPx) / autoRows : 0;
  const rows = table.rows.map((row, r) => ({
    heightPx: declaredHeights[r]! > 0 ? declaredHeights[r]! : autoHeightPx,
    cells: origins
      .filter((o) => o.row === r)
      .map(({ cell, col, spanW, spanH }) => {
        const rules = regionRulesAt(style, r, col, nRows, nCols);
        const fill = solidFillOf(cell.fill) ?? rules.fill;
        const opacity = fillOpacityOf(cell.fill);
        const edge = (key: "top" | "right" | "bottom" | "left") =>
          cell.borders?.[key]
            ? borderEdgeOf(cell.borders[key])
            : styleEdge(rules, key, r, col, spanW, spanH);
        const edges = {
          top: edge("top"),
          right: edge("right"),
          bottom: edge("bottom"),
          left: edge("left"),
        };
        const borders = Object.fromEntries(Object.entries(edges).filter(([, v]) => v));
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
    height: frameHeightPx,
    ...(childPath ? { childPath } : {}),
    table: { columnWidthsPx: widths, rows },
  };
}
