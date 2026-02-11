import { TextRun, ExternalHyperlink, type IParagraphOptions } from "docx";
import { convertTextNodes } from "./text";
import type { DetailsSummaryNode } from "@docen/extensions/types";
import type { DocxExportOptions } from "../options";

/**
 * Convert TipTap detailsSummary node to paragraph options
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * @param node - TipTap detailsSummary node
 * @param params - Conversion parameters
 * @returns Paragraph options (pure data object) with summary styling
 */
export function convertDetailsSummary(
  node: DetailsSummaryNode,
  params: {
    /** Export options for details styling */
    options?: DocxExportOptions["details"];
  },
): IParagraphOptions {
  // Convert summary content to text runs
  const summaryChildren = convertTextNodes(node.content || []).filter(
    (item): item is TextRun | ExternalHyperlink => item !== undefined,
  );

  // Return paragraph options with styling
  return {
    children: summaryChildren,
    ...params.options?.summary?.paragraph,
  };
}
