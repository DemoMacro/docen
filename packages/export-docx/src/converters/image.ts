import { ImageRun, IImageOptions, PositiveUniversalMeasure } from "docx-plus";
import { ImageNode } from "@docen/extensions/types";
import {
  getImageTypeFromSrc,
  getImageWidth,
  getImageHeight,
  defaultImageHandler,
  type DocxImageExportHandler,
} from "../utils";
import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";

/**
 * Convert TipTap image node to DOCX ImageRun
 *
 * @param node - TipTap image node
 * @param params - Conversion parameters
 * @returns Promise<DOCX ImageRun>
 */
export async function convertImage(
  node: ImageNode,
  params?: {
    /** Maximum available width (number = pixels, or string like "6in", "152.4mm") */
    maxWidth?: number | PositiveUniversalMeasure;
    /** Additional image options to apply */
    options?: Partial<IImageOptions>;
    /** Custom image handler for fetching image data */
    handler?: DocxImageExportHandler;
  },
): Promise<ImageRun> {
  // Get image data and metadata
  let imageData: Uint8Array;
  let imageMeta: ImageMeta;
  try {
    const src = node.attrs?.src || "";
    const handler = params?.handler ?? defaultImageHandler;
    imageData = await handler(src);

    // Extract metadata from image data
    try {
      imageMeta = getImageMetadata(imageData);
    } catch {
      imageMeta = {
        type: getImageTypeFromSrc(src),
        width: undefined,
        height: undefined,
        orientation: undefined,
      };
    }
  } catch (error) {
    console.warn(`Failed to process image:`, error);
    // Return placeholder ImageRun
    return new ImageRun({
      type: "png",
      data: new Uint8Array(0),
      transformation: { width: 100, height: 100 },
      altText: { name: node.attrs?.alt || "Failed to load image" },
    });
  }

  // Determine final dimensions: first from node attrs, then from image metadata
  const finalWidth = getImageWidth(node, imageMeta, params?.maxWidth);
  const finalHeight = getImageHeight(node, finalWidth, imageMeta, params?.maxWidth);

  // Build transformation object
  const transformation: {
    width: number;
    height: number;
    rotation?: number;
  } = {
    width: finalWidth,
    height: finalHeight,
  };

  // Add rotation if present (in degrees)
  if (node.attrs?.rotation !== undefined) {
    transformation.rotation = node.attrs.rotation;
  }

  const imageType = getImageTypeFromSrc(node.attrs?.src || "");

  // Build common options (shared between SVG and regular images)
  const commonOptions = {
    // Apply global options first
    ...params?.options,
    transformation,
    altText: {
      name: node.attrs?.alt || "",
      description: undefined,
      title: node.attrs?.title || undefined,
    },
    // Apply floating positioning from node.attrs if present
    ...(node.attrs?.floating && {
      floating: node.attrs.floating,
    }),
    // Apply outline from node.attrs if present
    ...(node.attrs?.outline && {
      outline: node.attrs.outline,
    }),
  };

  // SVG requires a fallback raster image
  if (imageType === "svg") {
    return new ImageRun({
      ...commonOptions,
      type: "svg",
      data: imageData,
      fallback: {
        type: "png",
        data: new Uint8Array(0), // Caller should provide proper fallback via handler
      },
    });
  }

  return new ImageRun({
    ...commonOptions,
    type: imageType,
    data: imageData,
    // Apply crop (srcRect) from node.attrs if present (only for regular images)
    ...(node.attrs?.crop && {
      srcRect: node.attrs.crop,
    }),
  });
}
