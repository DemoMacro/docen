import { Paragraph, IParagraphOptions } from "docx";
import { convertText, convertHardBreak } from "./text";
import { convertImageToRun } from "./image";
import { ParagraphNode, ImageNode } from "../types";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap paragraph node to DOCX Paragraph
 *
 * @param node - TipTap paragraph node
 * @param options - Optional paragraph options (e.g., numbering)
 * @param exportOptions - Export options (for image processing)
 * @returns Promise<DOCX Paragraph object>
 */
export async function convertParagraph(
  node: ParagraphNode,
  options?: IParagraphOptions,
  exportOptions?: DocxExportOptions,
): Promise<Paragraph> {
  // Convert content to text runs and images
  const children = await Promise.all(
    (node.content || []).map(async (contentNode) => {
      if (contentNode.type === "text") {
        return convertText(contentNode);
      } else if (contentNode.type === "hardBreak") {
        return convertHardBreak(contentNode.marks);
      } else if (contentNode.type === "image") {
        // Convert image node to ImageRun directly
        const imageRun = await convertImageToRun(contentNode as ImageNode, exportOptions?.image);
        return imageRun;
      }
      return [];
    }),
  );

  // Flatten the array of arrays
  const flattenedChildren = children.flat();

  // Create paragraph with options
  const paragraphOptions: IParagraphOptions = {
    children: flattenedChildren,
    ...options,
  };

  return new Paragraph(paragraphOptions);
}
