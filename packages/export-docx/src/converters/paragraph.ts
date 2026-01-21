import { Paragraph, IParagraphOptions, type IImageOptions } from "docx";
import { convertText, convertHardBreak } from "./text";
import { convertImage } from "./image";
import { ParagraphNode, ImageNode } from "@docen/extensions/types";
import { applyParagraphStyleAttributes } from "../utils";
import type { PositiveUniversalMeasure } from "docx";
import type { DocxImageExportHandler } from "../utils/image";

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
      /** Additional image options to apply */
      options?: Partial<IImageOptions>;
      /** Custom image handler for fetching image data */
      handler?: DocxImageExportHandler;
    };
  },
): Promise<Paragraph> {
  const { options, image } = params || {};

  // Convert content to text runs and images
  const children = [];

  for (const contentNode of node.content || []) {
    if (contentNode.type === "text") {
      children.push(convertText(contentNode));
    } else if (contentNode.type === "hardBreak") {
      children.push(convertHardBreak(contentNode.marks));
    } else if (contentNode.type === "image") {
      const imageRun = await convertImage(contentNode as ImageNode, {
        maxWidth: image?.maxWidth,
        options: image?.options,
        handler: image?.handler,
      });
      children.push(imageRun);
    }
  }

  // Determine paragraph options
  let paragraphOptions: IParagraphOptions = {
    children,
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
    paragraphOptions = applyParagraphStyleAttributes(paragraphOptions, node.attrs);
  }

  return new Paragraph(paragraphOptions);
}
