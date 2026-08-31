import type { FlowPage } from "@docen/layout";

/**
 * Structural equality over two plain-data subtrees. Dynamic key walk (not a
 * hand-listed field set) so a new field on any LaidOut* type is covered the
 * day it lands — a missed field here would freeze a page at stale pixels.
 * Identical references short-circuit (the projection's identity caches hand
 * back the same media strings, so unchanged images compare in O(1)).
 *
 * False negatives are impossible by construction (any difference walks to a
 * differing primitive); false positives — structurally equal subtrees that
 * were rebuilt — merely repaint a page that did not need it.
 */
export function deepEq(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) {
    // Primitives (and NaN — never equal, a rebuilt float would repaint) fall
    // out here; functions/undefined never appear on laid-out data.
    return false;
  }
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!deepEq(a[i], b[i])) return false;
    return true;
  }
  const ka = Object.keys(a as Record<string, unknown>);
  const kb = Object.keys(b as Record<string, unknown>);
  if (ka.length !== kb.length) return false;
  for (const k of ka) {
    if (!deepEq((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k])) {
      return false;
    }
  }
  return true;
}

/** Which pages of `next` differ from `prev`, position by position (a page
 *  past either end, or an empty `prev`, is dirty). Pages whose item
 *  placements — yPx plus the whole laid-out block tree — match keep their
 *  canvas; a mid-document shift marks everything from its offset on. */
export function dirtyPagesOf(
  prev: readonly FlowPage[] | undefined,
  next: readonly FlowPage[],
): boolean[] {
  const dirty = next.map(() => true);
  if (!prev) return dirty;
  const n = Math.min(prev.length, next.length);
  for (let i = 0; i < n; i++) {
    const pi = prev[i]!.items;
    const ni = next[i]!.items;
    let same = pi.length === ni.length;
    for (let j = 0; same && j < pi.length; j++) {
      same = pi[j]!.yPx === ni[j]!.yPx && deepEq(pi[j]!.block, ni[j]!.block);
    }
    dirty[i] = !same;
  }
  return dirty;
}
