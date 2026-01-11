import { Paragraph, BorderStyle } from "docx";
import { HorizontalRuleNode } from "../types";
import type { DocxExportOptions } from "../option";

/**
 * Convert TipTap horizontalRule node to DOCX Paragraph
 * Creates a horizontal line using bottom border
 *
 * @param node - TipTap horizontalRule node
 * @param options - Export options for horizontal rule styling
 * @returns DOCX Paragraph object with horizontal rule styling
 */
export function convertHorizontalRule(
  node: HorizontalRuleNode,
  options: DocxExportOptions["horizontalRule"],
): Paragraph {
  return new Paragraph({
    children: [], // Empty content
    border: {
      bottom: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "auto",
      },
    },
    ...options?.paragraph,
  });
}
