import type { JSONContent } from "@docen/docx/core";

import { mergeSplitParagraphs } from "./paragraph-split";
import { mergeSplitTables } from "./table-split";

/**
 * Page ↔ flat document shape conversion.
 *
 * The DOCX round-trip shape is flat `doc > block+` (what `parseDOCX` returns
 * and `generateDOCX` consumes). The editing-time shape is `doc > page+` (each
 * page a fixed-height box). These helpers bridge the two at the editor layer —
 * the docx converters are untouched, so page nodes never leak into DOCX.
 */

/** Wrap a flat `doc > block+` doc into `doc > page > block+` for editing. All
 *  blocks go into a single page; the paginator splits them across pages once
 *  measured. Used on DOCX import / initial content. */
export function wrapPages(json: JSONContent | undefined): JSONContent {
  const blocks = json?.content ?? [];
  // Already wrapped (doc > page+) — pass through unchanged. Lets callers
  // (e.g. setJSON) accept either flat or pre-wrapped JSON safely.
  if (blocks.length > 0 && blocks.every((c) => c.type === "page")) {
    return { ...json, type: "doc", content: blocks };
  }
  const pageContent = blocks.length ? blocks : [{ type: "paragraph" }];
  return {
    ...json,
    type: "doc",
    content: [{ type: "page", content: pageContent }],
  };
}

/** Unwrap `doc > page+` back to flat `doc > block+` for DOCX export. Page
 *  boundaries vanish — they carried no DOCX meaning. Any stray non-page
 *  top-level node (defensive) passes through. */
export function unwrapPages(json: JSONContent): JSONContent {
  const blocks: JSONContent[] = [];
  for (const child of json.content ?? []) {
    if (child.type === "page") blocks.push(...(child.content ?? []));
    else blocks.push(child);
  }
  // Merge paragraphs/tables split across pages back into single nodes (drop
  // editor-only split attrs) — the split is round-trip-transparent. Paragraphs
  // first (they wrap tables in the flow), then tables.
  return { ...json, type: "doc", content: mergeSplitTables(mergeSplitParagraphs(blocks)) };
}
