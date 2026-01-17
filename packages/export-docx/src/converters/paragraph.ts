import { Paragraph, IParagraphOptions } from "docx";
import { convertText, convertHardBreak } from "./text";
import { convertImage } from "./image";
import { ParagraphNode, ImageNode } from "../types";
import { convertCssLengthToPixels, convertPixelsToTwip } from "../utils";
import type { PositiveUniversalMeasure } from "docx";

/**
 * Convert TipTap paragraph node to DOCX Paragraph
 *
 * @param node - TipTap paragraph node
 * @param params - Conversion parameters
 * @returns Promise<DOCX Paragraph object>
 */
export async function convertParagraph(
  node: ParagraphNode,
  params?: {
    options?: IParagraphOptions;
    /** Image conversion parameters */
    image?: {
      /** Maximum available width for inline images (number = pixels, or string like "6in", "152.4mm") */
      maxWidth?: number | PositiveUniversalMeasure;
    };
  },
): Promise<Paragraph> {
  const { options, image } = params || {};

  // Convert content to text runs and images
  const children = await Promise.all(
    (node.content || []).map(async (contentNode) => {
      if (contentNode.type === "text") {
        return convertText(contentNode);
      } else if (contentNode.type === "hardBreak") {
        return convertHardBreak(contentNode.marks);
      } else if (contentNode.type === "image") {
        // Convert image node to ImageRun directly
        const imageRun = await convertImage(contentNode as ImageNode, {
          maxWidth: image?.maxWidth,
        });
        return imageRun;
      }
      return [];
    }),
  );

  // Flatten the array of arrays
  const flattenedChildren = children.flat();

  // Determine paragraph options
  let paragraphOptions: IParagraphOptions = {
    children: flattenedChildren,
  };

  // Apply any passed-in options (e.g., numbering for lists, style references)
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
          ...(indentLeft && { left: convertPixelsToTwip(convertCssLengthToPixels(indentLeft)) }),
          ...(indentRight && { right: convertPixelsToTwip(convertCssLengthToPixels(indentRight)) }),
          ...(indentFirstLine && {
            firstLine: convertPixelsToTwip(convertCssLengthToPixels(indentFirstLine)),
          }),
        },
      };
    }

    // Convert spacing to DOCX format
    if (spacingBefore || spacingAfter) {
      paragraphOptions = {
        ...paragraphOptions,
        spacing: {
          ...(spacingBefore && {
            before: convertPixelsToTwip(convertCssLengthToPixels(spacingBefore)),
          }),
          ...(spacingAfter && {
            after: convertPixelsToTwip(convertCssLengthToPixels(spacingAfter)),
          }),
        },
      };
    }
  }

  return new Paragraph(paragraphOptions);
}
