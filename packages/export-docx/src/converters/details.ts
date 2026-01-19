import { Paragraph, TextRun, ExternalHyperlink } from "docx";
import { convertTextNodes } from "./text";
import type { DetailsSummaryNode } from "@docen/extensions/types";
import type { DocxExportOptions } from "../options";

/**
 * Convert TipTap detailsSummary node to DOCX Paragraph
 *
 * @param node - TipTap detailsSummary node
 * @param params - Conversion parameters
 * @returns DOCX Paragraph with summary styling
 */
export function convertDetailsSummary(
  node: DetailsSummaryNode,
  params: {
    /** Export options for details styling */
    options?: DocxExportOptions["details"];
  },
): Paragraph {
  // Convert summary content to text runs
  const summaryChildren = convertTextNodes(node.content || []).filter(
    (item): item is TextRun | ExternalHyperlink => item !== undefined,
  );

  // Create summary paragraph with styling
  return new Paragraph({
    children: summaryChildren,
    ...params.options?.summary?.paragraph,
  });
}
