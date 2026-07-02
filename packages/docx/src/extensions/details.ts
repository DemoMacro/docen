import {
  Details as DetailsBase,
  DetailsSummary as DetailsSummaryBase,
  DetailsContent as DetailsContentBase,
} from "@tiptap/extension-details";

/**
 * Details extension — owns the DOCX expression of a collapsible details block.
 *
 * DOCX has no native collapsible region, but a block-level group-SDT is a
 * reversible container. The details maps to one group-SDT tagged "docen-
 * details"; the summary paragraph is marked with a fixed style so resolve can
 * split it back out from the content paragraphs. Structure round-trips fully
 * (summary + content); Word shows it expanded (no collapse) — an inherent
 * DOCX limitation, not data loss.
 */

/** SDT tag marking a details group content control. */
export const DETAILS_TAG = "docen-details";

/** Paragraph style marking the summary line within a details group-SDT. */
export const DETAILS_SUMMARY_STYLE = "DocenDetailsSummary";

// DocxManager assembles/parses the group-SDT using the constants above; the
// extensions carry no DOCX attrs of their own.
export {
  DetailsBase as Details,
  DetailsSummaryBase as DetailsSummary,
  DetailsContentBase as DetailsContent,
};
