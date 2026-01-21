import { ImageRun, IImageOptions, PositiveUniversalMeasure } from "docx";
import { ImageNode } from "@docen/extensions/types";
import {
  getImageTypeFromSrc,
  getImageWidth,
  getImageHeight,
  getImageDataAndMeta,
  convertToDocxImageType,
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
  // Get image type from metadata or URL
  const getImageType = (metaType?: string): "jpg" | "png" | "gif" | "bmp" => {
    // Try metadata type first, then fallback to URL-based type detection
    const typeKey = metaType || getImageTypeFromSrc(node.attrs?.src || "");
    return convertToDocxImageType(typeKey);
  };

  // Get image data and metadata
  let imageData: Uint8Array;
  let imageMeta: ImageMeta;
  try {
    const src = node.attrs?.src || "";

    // Use custom handler if provided
    if (params?.handler) {
      imageData = await params.handler(src);

      // Extract metadata from fetched data
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
    } else if (src.startsWith("http")) {
      // Default behavior: fetch HTTP images
      const result = await getImageDataAndMeta(src);
      imageData = result.data;
      imageMeta = result.meta;
    } else if (src.startsWith("data:")) {
      // Handle data URLs - extract the base64 part
      const base64Data = src.split(",")[1];

      if (!base64Data) {
        throw new Error("Invalid data URL: missing base64 data");
      }

      // Use TextEncoder to create Uint8Array from base64 (works in both Node and browser)
      const binaryString = atob(base64Data);
      const bytes = Uint8Array.from(binaryString, (char) => char.charCodeAt(0));
      imageData = bytes;

      // Extract metadata from data URL
      try {
        imageMeta = getImageMetadata(imageData);
      } catch {
        imageMeta = {
          type: "png",
          width: undefined,
          height: undefined,
          orientation: undefined,
        };
      }
    } else {
      throw new Error(`Unsupported image source format: ${src.substring(0, 20)}...`);
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
  // Note: docx library will handle the conversion to DOCX format (1/60000 degrees) internally
  if (node.attrs?.rotation !== undefined) {
    transformation.rotation = node.attrs.rotation;
  }

  // Build ImageRun options
  const imageOptions: IImageOptions = {
    // Apply global options first
    ...params?.options,
    // Required fields (override global options)
    type: getImageType(imageMeta.type),
    data: imageData,
    transformation,
    altText: {
      name: node.attrs?.alt || "",
      description: undefined,
      title: node.attrs?.title || undefined,
    },
    // Apply floating positioning from node.attrs if present
    ...(node.attrs?.floating && {
      floating: node.attrs.floating, // Type assertion needed for compatibility
    }),
    // Apply outline from node.attrs if present
    ...(node.attrs?.outline && {
      outline: node.attrs.outline, // Type assertion needed for compatibility
    }),
  };

  return new ImageRun(imageOptions);
}
