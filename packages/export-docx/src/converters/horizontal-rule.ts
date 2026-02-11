import { PageBreak, type IParagraphOptions } from "docx";
import { HorizontalRuleNode } from "@docen/extensions/types";
import type { DocxExportOptions } from "../options";

/**
 * Convert TipTap horizontalRule node to paragraph options
 *
 * This converter only handles data transformation from node to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * Uses page break by default (consistent with import-docx behavior)
 *
 * @param node - TipTap horizontalRule node
 * @param params - Conversion parameters
 * @returns Paragraph options (pure data object) with page break or custom styling
 */
export function convertHorizontalRule(
  node: HorizontalRuleNode,
  params: {
    /** Export options for horizontal rule styling */
    options?: DocxExportOptions["horizontalRule"];
  },
): IParagraphOptions {
  // Default: use page break (consistent with import-docx which detects page breaks as horizontal rules)
  const paragraphOptions: IParagraphOptions = {
    children: [new PageBreak()],
  };

  // Allow user to override with custom styling (e.g., border instead of page break)
  return {
    ...paragraphOptions,
    ...params.options?.paragraph,
  };
}
