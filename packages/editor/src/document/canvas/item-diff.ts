import type { FlowItem } from "@docen/layout";

import { deepEq } from "./page-eq";

/** One flow item's per-item repaint verdict: "keep" items are structurally
 *  unchanged (a deep-equal block at the same column x) and only translate
 *  vertically — their painted groups survive, so Leafer's pattern and image
 *  caches stay warm (the repaint hotspot the viewport work exposed);
 *  "repaint" items rebuild their group in place. */
export interface ItemOp {
  kind: "keep" | "repaint";
  /** The item's position in both arrays (position pairing — no LCS). */
  index: number;
}

/** Diff two generations of one page's flow items. A different item count (a
 *  paragraph added, removed, or split across pages) has no meaningful
 *  position pairing — the caller rebuilds the whole body. An insertion
 *  shifts later items down, but their blocks re-flow at the new width, so
 *  they diff as repaints — correct, just not minimal. */
export function diffFlowItems(
  prev: readonly FlowItem[] | undefined,
  next: readonly FlowItem[],
): ItemOp[] | null {
  if (!prev || prev.length !== next.length) return null;
  const ops: ItemOp[] = [];
  for (let i = 0; i < next.length; i++) {
    const p = prev[i]!;
    const n = next[i]!;
    const keep = (p.xPx ?? 0) === (n.xPx ?? 0) && deepEq(p.block, n.block);
    ops.push({ kind: keep ? "keep" : "repaint", index: i });
  }
  return ops;
}
