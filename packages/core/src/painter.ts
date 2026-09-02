/**
 * Scene painter — walks a laid-out page (the @docen/layout result) and builds
 * the LeaferJS tree. The layout engine owns ALL geometry; painting positions
 * what it is given and never measures. Text elements carry explicit width AND
 * height — Leafer never paints an element whose height is still 0.
 *
 * @module
 */
import type { FlowItem, LaidOutBlock, LaidOutStackItem } from "@docen/layout";
import { columnBoxesOf } from "@docen/layout";
import { Line, Rect, Text, type IGroup } from "leafer-ui";

import type { PaintColumn, PaintContext } from "./paint/context";
import { paintBreakRow, paintParagraph } from "./paint/paragraph";
import { paintTable } from "./paint/table";

export * from "./paint/context";
export * from "./paint/image";
export * from "./paint/drawing";
export * from "./paint/table";
export * from "./paint/paragraph";

export function paintScene(tree: IGroup, items: readonly FlowItem[], ctx: PaintContext): void {
  for (const item of items) {
    paintBlock(
      tree,
      item.block,
      ctx.flow.contentLeftPx + (item.xPx ?? 0),
      ctx.flow.contentTopPx + item.yPx,
      ctx,
    );
  }
}

/** Paint the section's column separator lines (w:cols/@w:sep) — one vertical
 *  line centered in each gap between neighboring columns, spanning the
 *  content box. */
export function paintColumnSeparators(tree: IGroup, ctx: PaintContext): void {
  const cols = ctx.columns;
  if (!cols?.separate || cols.count < 2) return;
  const boxes = columnBoxesOf(ctx.flow.contentWidthPx, cols);
  for (let i = 0; i < boxes.length - 1; i++) {
    const x = ctx.flow.contentLeftPx + boxes[i]!.xPx + boxes[i]!.widthPx + cols.spacePx / 2;
    tree.add(
      new Line({
        points: [x, ctx.flow.contentTopPx, x, ctx.flow.contentTopPx + ctx.flow.contentHeightPx],
        stroke: "#000000",
        strokeWidth: 1,
        hittable: false,
      }),
    );
  }
}

/** Paint this page's line numbers (w:lnNumType) in the left margin — each
 *  number's right edge sits `distancePx` left of the text margin and its box
 *  top aligns with the counted line's text box (same strut size, so the
 *  baselines agree). The marks arrive pre-counted from the stage; painting
 *  never measures beyond the label box. */
export function paintLineNumbers(tree: IGroup, ctx: PaintContext): void {
  const ln = ctx.lineNumbers;
  if (!ln || ln.marks.length === 0) return;
  const widest = Math.max(...ln.marks.map((m) => m.sizePx));
  const boxWidth = String(ln.marks[ln.marks.length - 1]!.num).length * widest * 0.62;
  for (const mark of ln.marks) {
    tree.add(
      new Text({
        x: ctx.flow.contentLeftPx - ln.config.distancePx - boxWidth,
        y: ctx.flow.contentTopPx + mark.yPx,
        width: boxWidth,
        height: mark.sizePx * 1.4,
        text: String(mark.num),
        fill: "#000000",
        fontSize: mark.sizePx,
        textAlign: "right",
      }),
    );
  }
}

/** Paint a pre-laid header/footer stack at its page position. */
export function paintFurnitureStack(
  tree: IGroup,
  stack: readonly LaidOutStackItem[],
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  for (const item of stack) {
    paintBlock(tree, item.block, x, y + item.yPx, ctx);
  }
}

export function paintBlock(
  tree: IGroup,
  block: LaidOutBlock,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): void {
  switch (block.kind) {
    case "paragraph":
      paintParagraph(tree, block, x, y, ctx, col);
      return;
    // Only paragraphs can carry drawings; the behind pass therefore skips
    // every other block so nothing paints twice.
    case "table":
    case "placeholder":
    case "pageBreak":
      if (ctx.layer === "behind") return;
      if (block.kind === "table") paintTable(tree, block, x, y, ctx);
      else if (block.kind === "placeholder") paintPlaceholder(tree, block, x, y);
      else if (ctx.showMarks) paintBreakRow(tree, block, x, y, ctx.flow.contentWidthPx, ctx);
      return;
    case "group":
      for (const child of block.children) {
        paintBlock(tree, child.block, x, y + child.yPx, ctx, col);
      }
      return;
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
