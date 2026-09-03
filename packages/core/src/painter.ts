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
  const cols = columnBoxesOf(ctx.flow.contentWidthPx, ctx.columns);
  for (const item of items) {
    const x = ctx.flow.contentLeftPx + (item.xPx ?? 0);
    paintBlock(
      tree,
      item.block,
      x,
      ctx.flow.contentTopPx + item.yPx,
      ctx,
      // A multi-column item paints within its column box: shading fills,
      // paragraph borders and the break rows' rules span the column, not the
      // whole content width (the table walk threads its cell width the same
      // way). The interval match survives float drift between the flow's
      // stamped x and the boxes recomputed here.
      item.xPx != null
        ? {
            width:
              cols.find((c) => item.xPx! >= c.xPx - 0.01 && item.xPx! < c.xPx + c.widthPx)
                ?.widthPx ?? cols[0]!.widthPx,
            inCell: false,
          }
        : undefined,
    );
  }
}

/** Paint the document grid (Word's View → Gridlines): one horizontal rule
 *  every `linePitchPx` across the content box — the pitch the docGrid snaps
 *  body lines to, so the overlay shows why lines sit where they sit. */
export function paintGridlines(tree: IGroup, ctx: PaintContext): void {
  const pitch = ctx.flow.linePitchPx;
  if (!ctx.showGridlines || !pitch || pitch <= 0) return;
  const left = ctx.flow.contentLeftPx;
  const right = left + ctx.flow.contentWidthPx;
  for (
    let y = ctx.flow.contentTopPx;
    y <= ctx.flow.contentTopPx + ctx.flow.contentHeightPx;
    y += pitch
  ) {
    tree.add(
      new Line({
        points: [left, y, right, y],
        stroke: "#c5d3ee",
        strokeWidth: 1,
        hittable: false,
      }),
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
 *  baselines agree). Auto placement (w:distance omitted, `distancePx: null` —
 *  the OOXML default) keeps a small fixed gap from the text margin — about
 *  the midpoint of the stage's crop-mark leg, where the number reads as
 *  beside the text, not stranded mid-margin.
 *  The marks arrive pre-counted from the stage; painting never measures
 *  beyond the label box. */
export function paintLineNumbers(tree: IGroup, ctx: PaintContext): void {
  const ln = ctx.lineNumbers;
  if (!ln || ln.marks.length === 0) return;
  const widest = Math.max(...ln.marks.map((m) => m.sizePx));
  const boxWidth = String(ln.marks[ln.marks.length - 1]!.num).length * widest * 0.62;
  const { distancePx } = ln.config;
  const labelX =
    distancePx != null
      ? ctx.flow.contentLeftPx - distancePx - boxWidth
      : Math.max(0, ctx.flow.contentLeftPx - 12 - boxWidth);
  for (const mark of ln.marks) {
    tree.add(
      new Text({
        x: labelX,
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
      else if (ctx.showMarks)
        paintBreakRow(tree, block, x, y, col?.width ?? ctx.flow.contentWidthPx, ctx);
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
