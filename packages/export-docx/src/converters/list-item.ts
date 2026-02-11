import { IParagraphOptions } from "docx";
import { ListItemNode } from "@docen/extensions/types";
import { convertParagraph } from "./paragraph";

/**
 * Convert TipTap list item node to paragraph options
 *
 * This converter only handles data transformation from node content to DOCX format properties.
 * It returns pure data objects (IParagraphOptions), not DOCX instances.
 *
 * Note: The numbering reference (including start value) is typically
 * handled by the parent list converter. This function focuses on
 * converting the paragraph content of the list item.
 *
 * @param node - TipTap list item node
 * @param params - Conversion parameters
 * @returns Promise<Paragraph options (pure data object)>
 */
export async function convertListItem(
  node: ListItemNode,
  params: {
    options?: IParagraphOptions;
  },
): Promise<IParagraphOptions> {
  if (!node.content || node.content.length === 0) {
    return {};
  }

  // Convert the first paragraph in the list item
  const firstParagraph = node.content[0];
  if (firstParagraph.type === "paragraph") {
    return await convertParagraph(firstParagraph, {
      options: params.options,
    });
  }

  // Fallback to empty paragraph options
  return {};
}
