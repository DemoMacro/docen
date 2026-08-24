/**
 * Scene painter — walks a laid-out page (the @docen/layout result) and builds
 * the LeaferJS tree. The layout engine owns ALL geometry; painting positions
 * what it is given and never measures. Text elements carry explicit width AND
 * height — Leafer never paints an element whose height is still 0.
 */
import {
  isCjkCodeUnit,
  type FlowItem,
  type LaidOutBlock,
  type LaidOutParagraph,
  type LaidOutTable,
  type LayoutInline,
  type LayoutTextStyle,
} from "@docen/layout";
import { Image as LeaferImage, Rect, Text, type IGroup } from "leafer-ui";

import type { CanvasStageContext } from "./stage";

/** Paint one page's flow items (block + page-content y). */
export function paintScene(
  tree: IGroup,
  items: readonly FlowItem[],
  ctx: CanvasStageContext,
): void {
  for (const item of items) {
    paintBlock(tree, item.block, ctx.flow.contentLeftPx, ctx.flow.contentTopPx + item.yPx, ctx);
  }
}

function paintBlock(
  tree: IGroup,
  block: LaidOutBlock,
  x: number,
  y: number,
  ctx: CanvasStageContext,
): void {
  switch (block.kind) {
    case "paragraph":
      paintParagraph(tree, block, x, y, ctx);
      return;
    case "table":
      paintTable(tree, block, x, y, ctx);
      return;
    case "group":
      for (const child of block.children) {
        paintBlock(tree, child.block, x, y + child.yPx, ctx);
      }
      return;
    case "placeholder":
      paintPlaceholder(tree, block, x, y);
      return;
    case "pageBreak":
      return;
  }
}

function paintParagraph(
  tree: IGroup,
  para: LaidOutParagraph,
  x: number,
  y: number,
  ctx: CanvasStageContext,
): void {
  // w:pBdr horizontal rules (the "Education" underline shape): between the
  // text and `spacePx` below it, spanning the wrapping width.
  for (const side of ["top", "bottom"] as const) {
    const edge = para.borders?.[side];
    if (!edge || !edge.px || edge.style === "nil" || edge.style === "none") continue;
    const top = side === "top";
    const lineY = top ? y - (edge.spacePx ?? 0) : y + para.heightPx + (edge.spacePx ?? 0);
    const right = ctx.flow.contentLeftPx + ctx.flow.contentWidthPx;
    tree.add(
      new Rect({
        x,
        y: lineY - (top ? 0 : edge.px),
        width: Math.max(0, right - x),
        height: edge.px,
        fill: "#1b1b1b",
      }),
    );
  }
  for (const [lineIndex, line] of para.lines.entries()) {
    const lineY = y + line.yPx;
    // Line x origin: the left indent, plus the first line's own indent —
    // item xPx starts at the line's content start (hanging indents go left).
    const lineX =
      x + (para.indent?.leftPx ?? 0) + (lineIndex === 0 ? (para.indent?.firstLinePx ?? 0) : 0);
    // A justified line stretches each text item to the next item's x (the
    // last one to the line's content width): Leafer's textAlign "both-letter"
    // spreads the slack as uniform letter spacing inside that interval —
    // Word's CJK justification model.
    const justified = line.justifyGapPx != null;
    const rights: number[] = [];
    if (justified) {
      let nextLeft = line.maxWidthPx ?? 0;
      for (let i = line.items.length - 1; i >= 0; i--) {
        rights[i] = nextLeft;
        nextLeft = line.items[i].xPx;
      }
    }
    for (const [itemIndex, item] of line.items.entries()) {
      const inline: LayoutInline | undefined = para.inline[item.inlineIndex];
      if (!inline) continue;
      if (item.kind === "text" && inline.kind === "text") {
        // Half-leading: center the font's natural box in the line box (the
        // CSS/DOM model), so baselines line up with the DOM route.
        const family = familyOf(inline.style, item.text);
        const boxPx =
          inline.style.sizePx *
          ctx.metrics.normalRatio({ family, bold: inline.style.bold === true });
        const pad = Math.max(0, (line.heightPx - boxPx) / 2);
        const intervalPx = justified ? rights[itemIndex] - item.xPx : undefined;
        tree.add(
          new Text({
            x: lineX + item.xPx,
            y: lineY + pad,
            // width ONLY on justified items (their stretch interval): a width
            // on every line would let Leafer wrap the slice again with its
            // own metrics (a phantom second line). textWrap "none" keeps the
            // interval from wrapping; height keeps the element paintable
            // (height 0 is skipped by Leafer).
            width: intervalPx,
            textWrap: justified ? "none" : undefined,
            textAlign: justified ? "both-letter" : undefined,
            height: Math.max(1, line.heightPx),
            text: item.text,
            fill: "#1b1b1b",
            fontFamily: family,
            fontSize: inline.style.sizePx,
            fontWeight: inline.style.bold ? "700" : "400",
            fontStyle: inline.style.italic ? "italic" : "normal",
            letterSpacing: inline.style.letterSpacingPx
              ? { type: "px", value: inline.style.letterSpacingPx }
              : undefined,
          }),
        );
      } else if (item.kind === "picture" && inline.kind === "picture") {
        if (inline.src) {
          tree.add(
            new LeaferImage({
              url: inline.src,
              x: lineX + item.xPx,
              y: lineY,
              width: item.widthPx,
              height: item.heightPx,
            }),
          );
        } else {
          // Linked-only picture (no bytes in the package): an empty frame.
          tree.add(
            new Rect({
              x: lineX + item.xPx,
              y: lineY,
              width: item.widthPx,
              height: item.heightPx,
              fill: "#f3f3f3",
              stroke: "#c4c4c4",
              strokeWidth: 1,
            }),
          );
        }
      }
    }
  }
}

function paintTable(
  tree: IGroup,
  table: LaidOutTable,
  x: number,
  y: number,
  ctx: CanvasStageContext,
): void {
  let rowY = y;
  for (const row of table.rows) {
    let col = 0;
    for (const cell of row.cells) {
      const cellX = x + table.columnWidthsPx.slice(0, col).reduce((a, b) => a + b, 0);
      const cellW = table.columnWidthsPx.slice(col, col + cell.colspan).reduce((a, b) => a + b, 0);
      // Hairline prototype borders; collapse-aware styling reads the mirrored
      // `borders` when fidelity lands.
      tree.add(
        new Rect({
          x: cellX,
          y: rowY,
          width: cellW,
          height: row.heightPx,
          stroke: "#c4c4c4",
          strokeWidth: 1,
        }),
      );
      const contentX = cellX + (cell.insets.left ?? 0);
      const contentY = rowY + (cell.insets.top ?? 0);
      for (const stacked of cell.stack) {
        paintBlock(tree, stacked.block, contentX, contentY + stacked.yPx, ctx);
      }
      col += cell.colspan;
    }
    rowY += row.heightPx;
  }
}

function paintPlaceholder(
  tree: IGroup,
  block: { heightPx: number; label?: string },
  x: number,
  y: number,
): void {
  const width = 240;
  const height = Math.max(20, block.heightPx);
  tree.add(
    new Rect({
      x,
      y,
      width,
      height,
      fill: "#fafafa",
      stroke: "#d0d0d0",
      strokeWidth: 1,
      dashPattern: [4, 3],
    }),
  );
  if (block.label) {
    tree.add(
      new Text({
        x: x + 8,
        y: y + 4,
        text: `${block.label} (not rendered yet)`,
        fill: "#9a9a9a",
        fontFamily: "Inter, sans-serif",
        fontSize: 11,
      }),
    );
  }
}

/** The font family a text slice paints in: a slot-aware pick over the item's
 *  own text (the engine already itemized runs; CJK slices take the eastAsia
 *  slot — the resolveFontFamily rule, applied at paint time). */
function familyOf(style: LayoutTextStyle, text: string): string {
  if (typeof style.family === "string") return style.family;
  const slots = style.family;
  if (text && isCjkCodeUnit(text, 0)) {
    return slots.eastAsia ?? slots.latin ?? "Inter";
  }
  return slots.latin ?? "Inter";
}
