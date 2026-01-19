import { Paragraph, TextRun, ExternalHyperlink } from "docx";
import { convertTextNodes } from "./text";
import { convertNode } from "../generator";
import { DetailsNode, DetailsSummaryNode, DetailsContentNode } from "../types";
import type { DocxExportOptions } from "../option";
import { calculateEffectiveContentWidth } from "../utils";

/**
 * Convert TipTap details node to array of DOCX Paragraphs
 * Simulates collapsible content using indentation and borders
 *
 * @param node - TipTap details node
 * @param params - Conversion parameters
 * @returns Array of DOCX Paragraph objects
 */
export async function convertDetails(
  node: DetailsNode,
  params: {
    /** Export options for details styling */
    options?: DocxExportOptions["details"];
    /** Full export options for nested content conversion */
    exportOptions: DocxExportOptions;
  },
): Promise<Paragraph[]> {
  if (!node.content) return [];

  const result: Paragraph[] = [];
  let summaryNode: DetailsSummaryNode | undefined;
  let contentNode: DetailsContentNode | undefined;

  // Find summary and content nodes
  for (const child of node.content) {
    if (child.type === "detailsSummary") {
      summaryNode = child;
    } else if (child.type === "detailsContent") {
      contentNode = child;
    }
  }

  // Convert summary (summary-style paragraph with border)
  if (summaryNode?.content) {
    const summaryChildren = convertTextNodes(summaryNode.content).filter(
      (item): item is TextRun | ExternalHyperlink => item !== undefined,
    );

    const summaryParagraph = new Paragraph({
      children: summaryChildren,
      ...params.options?.summary?.paragraph,
    });

    result.push(summaryParagraph);
  }

  // Convert content (indented paragraphs and other elements)
  if (contentNode?.content) {
    const effectiveContentWidth = calculateEffectiveContentWidth(params.exportOptions);

    for (const contentElement of contentNode.content) {
      const element = await convertNode(
        contentElement,
        params.exportOptions,
        effectiveContentWidth,
      );
      if (Array.isArray(element)) {
        result.push(...(element as Paragraph[]));
      } else if (element) {
        result.push(element as Paragraph);
      }
    }
  }

  return result;
}
