import type { ParagraphOptions, SectionChild } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import {
  Details as DetailsBase,
  DetailsSummary as DetailsSummaryBase,
  DetailsContent as DetailsContentBase,
} from "@tiptap/extension-details";

import type { ParseBlockRule, ResolveContext } from "./types";

/**
 * Details extension — owns the DOCX expression of a collapsible details block.
 *
 * DOCX has no native collapsible region, but a block-level group-SDT is a
 * reversible container. The details maps to one group-SDT tagged "docen-
 * details"; the summary paragraph is marked with a fixed style so resolve can
 * split it back out from the content paragraphs. Structure round-trips fully
 * (summary + content); Word shows it expanded (no collapse) — an inherent
 * DOCX limitation, not data loss.
 *
 * The `parseDocxBlock` rule (below) recognizes a details group-SDT during
 * resolve and rebuilds the details/detailsSummary/detailsContent nodes; the
 * extensions themselves carry no DOCX attrs of their own.
 */

/** SDT tag marking a details group content control. */
export const DETAILS_TAG = "docen-details";

/** Paragraph style marking the summary line within a details group-SDT. */
export const DETAILS_SUMMARY_STYLE = "DocenDetailsSummary";

// ── Block parse rule (resolve: group-SDT → details node) ──

/**
 * Declarative block parse rule: recognize a group-SDT tagged "docen-details"
 * and rebuild it as a details node (summary + content). DocxManager dispatches
 * every SectionChild through this rule before the paragraph/passthrough
 * fallbacks; a non-details SDT falls through to passthrough. */
export const parseDocxBlock: ParseBlockRule = {
  match: (child) =>
    "sdt" in child &&
    (child as { sdt?: { properties?: { tag?: string } } }).sdt?.properties?.tag === DETAILS_TAG,
  convert: (child, ctx) =>
    resolveDetailsSdt(
      (child as { sdt: { properties?: { tag?: string }; children?: SectionChild[] } }).sdt,
      ctx,
    ),
};

/** Resolve a details group-SDT: the summary-style paragraph becomes
 *  detailsSummary, the remaining blocks fold into detailsContent. */
function resolveDetailsSdt(
  sdt: { properties?: { tag?: string }; children?: SectionChild[] },
  ctx: ResolveContext,
): JSONContent {
  const content: JSONContent[] = [];
  let summary: JSONContent[] | null = null;
  for (const child of sdt.children ?? []) {
    if ("paragraph" in child) {
      const para = child.paragraph as ParagraphOptions;
      if ((para as unknown as Record<string, unknown>).style === DETAILS_SUMMARY_STYLE) {
        summary = ctx.resolveInlineContent(para);
        continue;
      }
    }
    const node = ctx.resolveBlock(child);
    if (!node) continue;
    if (Array.isArray(node)) content.push(...node);
    else content.push(node);
  }
  const details: JSONContent = { type: "details", content: [] };
  if (summary !== null) details.content!.push({ type: "detailsSummary", content: summary });
  if (content.length > 0) details.content!.push({ type: "detailsContent", content });
  return details;
}

export { DetailsSummaryBase as DetailsSummary, DetailsContentBase as DetailsContent };

export const Details = DetailsBase.extend({ parseDocxBlock });
