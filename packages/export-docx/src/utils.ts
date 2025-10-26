/**
 * Shared utility functions for the export-docx package
 */

import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";
import { ofetch } from "ofetch";

/**
 * Extract image type from URL or base64 data
 */
export function getImageTypeFromSrc(
  src: string,
): "png" | "jpeg" | "gif" | "bmp" | "tiff" {
  if (src.startsWith("data:")) {
    const match = src.match(/data:image\/(\w+);/);
    if (match) {
      const type = match[1].toLowerCase();
      switch (type) {
        case "jpg":
        case "jpeg":
          return "jpeg";
        case "png":
          return "png";
        case "gif":
          return "gif";
        case "bmp":
          return "bmp";
        case "tiff":
          return "tiff";
        default:
          return "png";
      }
    }
  } else {
    const extension = src.split(".").pop()?.toLowerCase();
    switch (extension) {
      case "jpg":
      case "jpeg":
        return "jpeg";
      case "png":
        return "png";
      case "gif":
        return "gif";
      case "bmp":
        return "bmp";
      case "tiff":
        return "tiff";
      default:
        return "png";
    }
  }
  return "png";
}

/**
 * Create floating options for full-width images
 */
export function createFloatingOptions() {
  return {
    horizontalPosition: {
      relative: "page",
      align: "center",
    },
    verticalPosition: {
      relative: "page",
      align: "top",
    },
    lockAnchor: true,
    behindDocument: false,
    inFrontOfText: false,
  };
}

/**
 * Get image width with clear priority: node attrs > options.run > image metadata > default
 */
export function getImageWidth(
  node: { attrs?: { width?: number | null } },
  options?: { run?: { transformation?: { width?: number } } },
  imageMeta?: { width?: number },
): number {
  if (node.attrs?.width) return node.attrs.width;
  if (options?.run?.transformation?.width)
    return options.run.transformation.width;
  if (imageMeta?.width) return Math.min(imageMeta.width, 600);
  return 400;
}

/**
 * Get image height with clear priority: node attrs > options.run > calculated > default
 */
export function getImageHeight(
  node: { attrs?: { height?: number | null } },
  width: number,
  options?: { run?: { transformation?: { height?: number } } },
  imageMeta?: { width?: number; height?: number },
): number {
  if (node.attrs?.height) return node.attrs.height;
  if (options?.run?.transformation?.height)
    return options.run.transformation.height;
  if (imageMeta?.width && imageMeta?.height)
    return Math.round((width * imageMeta.height) / imageMeta.width);
  return 300;
}

/**
 * Fetch image data and metadata from URL
 */
export async function getImageDataAndMeta(
  url: string,
): Promise<{ data: Uint8Array; meta: ImageMeta }> {
  try {
    // Use ofetch to get binary data with responseType: "blob"
    const blob = await ofetch(url, { responseType: "blob" });
    const data = await blob.bytes();

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
