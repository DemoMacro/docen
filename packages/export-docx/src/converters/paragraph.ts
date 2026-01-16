import { Paragraph, IParagraphOptions } from "docx";
import { convertText, convertHardBreak } from "./text";
import { convertImageToRun } from "./image";
import { ParagraphNode, ImageNode } from "../types";
import { DocxExportOptions } from "../option";

/**
 * Convert pixels to TWIPs (Twentieth of a Point)
 * 1 inch = 1440 TWIPs, 1px ≈ 15 TWIPs (at 96 DPI: 1px = 0.75pt = 15 TWIP)
 */
function pxToTwip(px: number): number {
  return Math.round(px * 15);
}

/**
 * Convert TipTap paragraph node to DOCX Paragraph
 *
 * @param node - TipTap paragraph node
 * @param params - Conversion parameters
 * @returns Promise<DOCX Paragraph object>
 */
export async function convertParagraph(
  node: ParagraphNode,
  params: {
    options?: IParagraphOptions;
    exportOptions?: DocxExportOptions;
  },
): Promise<Paragraph> {
  const { options, exportOptions } = params;

  // Check if paragraph contains only images (for image-specific styling)
  const onlyContainsImages =
    node.content &&
    node.content.length > 0 &&
    node.content.every((child) => child.type === "image");

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

  // Determine paragraph options based on content type
  let paragraphOptions: IParagraphOptions = {
    children: flattenedChildren,
  };

  // Apply default paragraph styles directly to each paragraph for better compatibility
  // instead of relying on <w:pPrDefault> which may not be handled correctly by some word processors
  if (!onlyContainsImages && exportOptions?.styles?.default?.document?.paragraph) {
    paragraphOptions = {
      ...paragraphOptions,
      ...exportOptions.styles.default.document.paragraph,
    };
  }

  if (onlyContainsImages && exportOptions?.image?.paragraph) {
    // Apply image-specific paragraph options (e.g., alignment for centering)
    paragraphOptions = {
      ...paragraphOptions,
      ...exportOptions.image.paragraph,
    };
  }

  // Apply any passed-in options (e.g., numbering for lists)
  if (options) {
    paragraphOptions = {
      ...paragraphOptions,
      ...options,
    };
  }

  // Handle paragraph style attributes from node.attrs
  if (node.attrs) {
    const { indentLeft, indentRight, indentFirstLine, spacingBefore, spacingAfter } = node.attrs;

    // Convert indentation to DOCX format
    if (indentLeft || indentRight || indentFirstLine) {
      paragraphOptions = {
        ...paragraphOptions,
        indent: {
          ...(indentLeft && { left: pxToTwip(indentLeft) }),
          ...(indentRight && { right: pxToTwip(indentRight) }),
          ...(indentFirstLine && { firstLine: pxToTwip(indentFirstLine) }),
        },
      };
    }

    // Convert spacing to DOCX format
    if (spacingBefore || spacingAfter) {
      paragraphOptions = {
        ...paragraphOptions,
        spacing: {
          ...(spacingBefore && { before: pxToTwip(spacingBefore) }),
          ...(spacingAfter && { after: pxToTwip(spacingAfter) }),
        },
      };
    }
  }

  return new Paragraph(paragraphOptions);
}
