import { Paragraph, ImageRun, IImageOptions } from "docx";
import { ImageNode } from "../types";
import { getImageTypeFromSrc } from "../utils";
import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";
import { PropertiesOptions } from "../option";

/**
 * Fetch image data and metadata from URL
 */
async function fetchImageData(
  url: string,
): Promise<{ data: Uint8Array; meta: ImageMeta }> {
  try {
    // For binary data, use fetch API directly following official pattern
    const fetchResponse = await fetch(url);
    const blob = await fetchResponse.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    // Get image metadata using image-meta
    let meta: ImageMeta;
    try {
      meta = getImageMetadata(data);
    } catch (error) {
      // If metadata extraction fails, use default values
      console.warn(`Failed to extract image metadata:`, error);
      meta = {
        width: undefined,
        height: undefined,
        type: getImageTypeFromSrc(url) || "png",
        orientation: undefined,
      };
    }

    return { data, meta };
  } catch (error) {
    console.warn(`Failed to fetch image from ${url}:`, error);
    throw error;
  }
}

/**
 * Convert TipTap image node to DOCX Paragraph with ImageRun
 *
 * @param node - TipTap image node
 * @param options - Image options from PropertiesOptions
 * @returns Promise<DOCX Paragraph> object with image
 */
export async function convertImage(
  node: ImageNode,
  options: PropertiesOptions["image"],
): Promise<Paragraph> {
  // Get image type from metadata or URL
  const getImageType = (metaType?: string): "jpg" | "png" | "gif" | "bmp" => {
    // Try metadata type first
    switch (metaType) {
      case "jpeg":
      case "jpg":
        return "jpg";
      case "png":
        return "png";
      case "gif":
        return "gif";
      case "bmp":
        return "bmp";
    }

    // Fallback to URL-based type detection
    const type = getImageTypeFromSrc(node.attrs?.src || "");
    switch (type) {
      case "jpeg":
        return "jpg";
      case "png":
        return "png";
      case "gif":
        return "gif";
      case "bmp":
        return "bmp";
      default:
        return "png";
    }
  };

  // Get image data and metadata
  let imageData: Uint8Array;
  let imageMeta: ImageMeta;
  try {
    const src = node.attrs?.src || "";
    if (src.startsWith("http")) {
      const result = await fetchImageData(src);
      imageData = result.data;
      imageMeta = result.meta;
    } else if (src.startsWith("data:")) {
      // Handle data URLs - extract the base64 part
      const base64Data = src.split(",")[1];
      // Use TextEncoder to create Uint8Array from base64 (works in both Node and browser)
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
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
      throw new Error(
        `Unsupported image source format: ${src.substring(0, 20)}...`,
      );
    }
  } catch (error) {
    console.warn(`Failed to process image:`, error);
    // Create placeholder paragraph
    return new Paragraph({
      children: [
        new ImageRun({
          type: "png",
          data: new Uint8Array(0), // Empty Uint8Array as placeholder
          transformation: { width: 100, height: 100 },
          altText: { name: node.attrs?.alt || "Failed to load image" },
        }),
      ],
    });
  }

  // Determine final dimensions
  // Priority: 1. Node attributes > 2. Image metadata > 3. Default size
  let finalWidth: number;
  let finalHeight: number;

  const nodeWidth = node.attrs?.width;
  const nodeHeight = node.attrs?.height;

  if (
    nodeWidth !== null &&
    nodeWidth !== undefined &&
    nodeHeight !== null &&
    nodeHeight !== undefined
  ) {
    // Use dimensions from node attributes
    finalWidth = nodeWidth;
    finalHeight = nodeHeight;
  } else if (imageMeta.width && imageMeta.height) {
    // Use dimensions from image metadata
    if (nodeWidth !== null && nodeWidth !== undefined) {
      // Scale height based on aspect ratio
      finalWidth = nodeWidth;
      finalHeight = Math.round(
        (nodeWidth * imageMeta.height) / imageMeta.width,
      );
    } else if (nodeHeight !== null && nodeHeight !== undefined) {
      // Scale width based on aspect ratio
      finalHeight = nodeHeight;
      finalWidth = Math.round(
        (nodeHeight * imageMeta.width) / imageMeta.height,
      );
    } else {
      // Use original dimensions, but limit if too large
      const maxSize = 600;
      if (imageMeta.width > maxSize || imageMeta.height > maxSize) {
        const scale = Math.min(
          maxSize / imageMeta.width,
          maxSize / imageMeta.height,
        );
        finalWidth = Math.round(imageMeta.width * scale);
        finalHeight = Math.round(imageMeta.height * scale);
      } else {
        finalWidth = imageMeta.width;
        finalHeight = imageMeta.height;
      }
    }
  } else {
    // Fallback to default size or optionsed default
    finalWidth = 400;
    finalHeight = 300;
  }

  // Build complete image options
  const imageOptions: IImageOptions = {
    type: getImageType(imageMeta.type),
    data: imageData,
    transformation: {
      width: finalWidth,
      height: finalHeight,
    },
    altText: {
      name: node.attrs?.alt || "",
      description: undefined,
      title: node.attrs?.title || undefined,
    },
  };

  // Create image run
  const imageRun = new ImageRun(imageOptions);

  // Create paragraph with image
  const paragraphOptions = options?.paragraph || {};
  const paragraph = new Paragraph({
    children: [imageRun],
    alignment: "center", // Default center align
    ...paragraphOptions, // Allow override of paragraph properties
  });

  return paragraph;
}
