import type { PositiveUniversalMeasure } from "docx";
import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";
import { ofetch } from "ofetch";
import { convertMeasureToPixels } from "./conversion";

const DEFAULT_MAX_IMAGE_WIDTH_PIXELS = 6.5 * 96; // A4 effective width in pixels

const MIME_TO_TYPE: Record<string, "png" | "jpeg" | "gif" | "bmp" | "tiff"> = {
  jpg: "jpeg",
  jpeg: "jpeg",
  png: "png",
  gif: "gif",
  bmp: "bmp",
  tiff: "tiff",
};

const EXT_TO_TYPE: Record<string, "png" | "jpeg" | "gif" | "bmp" | "tiff"> = {
  jpg: "jpeg",
  jpeg: "jpeg",
  png: "png",
  gif: "gif",
  bmp: "bmp",
  tiff: "tiff",
};

/**
 * Extract image type from URL or base64 data
 */

export const getImageTypeFromSrc = (src: string): "png" | "jpeg" | "gif" | "bmp" | "tiff" => {
  if (src.startsWith("data:")) {
    const match = src.match(/data:image\/(\w+);/);
    if (match) {
      const type = match[1].toLowerCase();
      return MIME_TO_TYPE[type] || "png";
    }
  } else {
    const extension = src.split(".").pop()?.toLowerCase();
    if (extension) {
      return EXT_TO_TYPE[extension] || "png";
    }
  }

  return "png";
};

/**
 * Calculate appropriate display size for image (mimicking Word's behavior)
 */

const calculateDisplaySize = (
  imageMeta: { width?: number; height?: number },
  maxWidthPixels: number = DEFAULT_MAX_IMAGE_WIDTH_PIXELS,
): { width: number; height: number } => {
  if (!imageMeta.width || !imageMeta.height) {
    return {
      width: maxWidthPixels,
      height: Math.round(maxWidthPixels * 0.75),
    };
  }

  if (imageMeta.width <= maxWidthPixels) {
    return {
      width: imageMeta.width,
      height: imageMeta.height,
    };
  }

  const scaleFactor = maxWidthPixels / imageMeta.width;
  return {
    width: maxWidthPixels,
    height: Math.round(imageMeta.height * scaleFactor),
  };
};

/**
 * Create floating options for full-width images
 */

export const createFloatingOptions = () => {
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
};

/**
 * Get image width with priority: node attrs > image meta > calculated > default
 */

export const getImageWidth = (
  node: { attrs?: { width?: number | null } },
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number => {
  if (node.attrs?.width !== undefined && node.attrs?.width !== null) {
    return node.attrs.width;
  }

  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.width;
  }

  return maxWidthPixels || DEFAULT_MAX_IMAGE_WIDTH_PIXELS;
};

/**
 * Get image height with priority: node attrs > image meta > calculated > default
 */

export const getImageHeight = (
  node: { attrs?: { height?: number | null } },
  width: number,
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number => {
  if (node.attrs?.height !== undefined && node.attrs?.height !== null) {
    return node.attrs.height;
  }

  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.height;
  }

  return Math.round(width * 0.75);
};

/**
 * Fetch image data and metadata from URL
 */

export const getImageDataAndMeta = async (
  url: string,
): Promise<{ data: Uint8Array; meta: ImageMeta }> => {
  try {
    const blob = await ofetch(url, { responseType: "blob" });
    const data = await blob.bytes();

    let meta: ImageMeta;
    try {
      meta = getImageMetadata(data);
    } catch (error) {
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
};
