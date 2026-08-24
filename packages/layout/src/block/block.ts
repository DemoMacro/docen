// Block dispatch + BFC stacking. The stacker is the vertical-margin model a
// table cell uses (ported from the editor's measureRowHeight): the first
// paragraph's `before` and the last `after` are contained (a BFC eats its
// edge margins), adjacent siblings collapse at the max of after/before.

import type { LayoutBlock, LayoutBlockContext } from "../layout-doc";
import type { LaidOutBlock, LaidOutStackItem } from "../layout-result";
import type { TextMeasurer } from "../text/measure";
import { layoutParagraph } from "./paragraph";
import { layoutTable } from "./table";

/** Lay out any block at `width` (the container's content width). */
export function layoutBlock(
  block: LayoutBlock,
  width: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): LaidOutBlock {
  switch (block.kind) {
    case "paragraph":
      return layoutParagraph(block, width, ctx, measurer);
    case "table":
      return layoutTable(block, width, ctx, measurer);
    case "group": {
      const stacked = stackBlocks(block.blocks, width, ctx, measurer);
      return { kind: "group", heightPx: stacked.heightPx, children: stacked.stack };
    }
    case "placeholder":
      return { kind: "placeholder", heightPx: block.heightPx, label: block.label };
    case "pageBreak":
      return { kind: "pageBreak", heightPx: 0 };
  }
}

export interface StackedBlocks {
  stack: LaidOutStackItem[];
  heightPx: number;
}

/** Stack sibling blocks vertically with collapsing paragraph margins (the
 *  cell/BFC model: edge margins count, middles collapse at the max). */
export function stackBlocks(
  blocks: readonly LayoutBlock[],
  width: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): StackedBlocks {
  const stack: LaidOutStackItem[] = [];
  let heightPx = 0;
  let prevAfter = 0;
  let first = true;
  for (const block of blocks) {
    const out = layoutBlock(block, width, ctx, measurer);
    const before = out.kind === "paragraph" ? out.beforePx : 0;
    const after = out.kind === "paragraph" ? out.afterPx : 0;
    const gap = first ? before : Math.max(prevAfter, before);
    heightPx += gap + out.heightPx;
    stack.push({ yPx: heightPx - out.heightPx, block: out });
    prevAfter = after;
    first = false;
  }
  heightPx += prevAfter; // the last paragraph's after is contained too
  return { stack, heightPx };
}
