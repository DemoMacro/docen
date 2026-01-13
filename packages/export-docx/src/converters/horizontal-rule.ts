import { Paragraph, PageBreak, IParagraphOptions } from "docx";
import { HorizontalRuleNode } from "../types";
import type { DocxExportOptions } from "../option";

/**
 * Convert TipTap horizontalRule node to DOCX Paragraph
 * Uses page break by default (consistent with import-docx behavior)
 *
 * @param node - TipTap horizontalRule node
 * @param options - Export options for horizontal rule styling
 * @returns DOCX Paragraph object with page break or custom styling
 */
export function convertHorizontalRule(
  node: HorizontalRuleNode,
  options: DocxExportOptions["horizontalRule"],
): Paragraph {
  // Default: use page break (consistent with import-docx which detects page breaks as horizontal rules)
  const paragraphOptions: IParagraphOptions = {
    children: [new PageBreak()],
  };

  // Allow user to override with custom styling (e.g., border instead of page break)
  return new Paragraph({
    ...paragraphOptions,
    ...options?.paragraph,
  });
}
