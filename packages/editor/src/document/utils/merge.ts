import type { JSONContent } from "@docen/docx/core";

/**
 * Merge helpers for C-route pagination — stitch editor-only splits back into
 * single nodes on export, so the page-level physical split is round-trip-
 * transparent (DOCX never sees `splitGroup`/`splitClone`/`splitPart`).
 *
 * Pure JSONContent transforms — no Tiptap/PM dependency. Lives in `utils/`
 * alongside `wrap.ts` (which calls these on `unwrapPages`), keeping the logic
 * layer independent of the extension layer (extensions → utils, one way).
 */

/** Merge adjacent paragraphs/headings that share a `splitGroup` back into one:
 *  concatenate inline content, clear the editor-only split attrs. Spacing is
 *  left untouched — the split kept the original spacing on both halves (only
 *  the page's final-paragraph `after` is clipped at layout time, not persisted),
 *  so merging restores the paragraph exactly as it was before the split.
 *  Mirrors `mergeSplitTables`. Used by `unwrapPages` on export. */
export function mergeSplitParagraphs(blocks: JSONContent[]): JSONContent[] {
  const merged: JSONContent[] = [];
  for (const block of blocks) {
    const last = merged[merged.length - 1];
    if (
      (block.type === "paragraph" || block.type === "heading") &&
      (last?.type === "paragraph" || last?.type === "heading") &&
      last.attrs?.splitGroup != null &&
      last.attrs?.splitGroup === block.attrs?.splitGroup
    ) {
      last.content = [...(last.content ?? []), ...(block.content ?? [])];
    } else {
      merged.push(block);
    }
  }
  for (const block of merged) {
    if ((block.type === "paragraph" || block.type === "heading") && block.attrs) {
      delete block.attrs.splitGroup;
      delete block.attrs.splitPart;
    }
  }
  return merged;
}

/** Merge adjacent table nodes that share a `splitGroup` back into one table:
 *  concatenate rows, drop `splitClone` continuation-header clones, and clear
 *  the editor-only split attrs. Used by `unwrapPages` on export so the split
 *  stays round-trip-transparent (the result is a single table per group). */
export function mergeSplitTables(blocks: JSONContent[]): JSONContent[] {
  const merged: JSONContent[] = [];
  for (const block of blocks) {
    const last = merged[merged.length - 1];
    if (
      block.type === "table" &&
      last?.type === "table" &&
      block.attrs?.splitGroup != null &&
      last.attrs?.splitGroup === block.attrs.splitGroup
    ) {
      const keptRows = (block.content ?? []).filter(
        (r) => r.type === "tableRow" && r.attrs?.splitClone !== true,
      );
      last.content = [...(last.content ?? []), ...keptRows];
    } else {
      merged.push(block);
    }
  }
  // Clear editor-only split attrs on every table (whole or merged).
  for (const block of merged) {
    if (block.type !== "table") continue;
    if (block.attrs) delete block.attrs.splitGroup;
    for (const row of block.content ?? []) {
      if (row.type === "tableRow" && row.attrs) delete row.attrs.splitClone;
    }
  }
  return merged;
}
