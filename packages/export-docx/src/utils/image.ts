import type { PositiveUniversalMeasure } from "docx-plus";
import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";
import { ofetch } from "ofetch";
import { convertMeasureToPixels, DOCX_DPI } from "@docen/utils";

// Custom image handler for fetching image data
export type DocxImageExportHandler = (src: string) => Promise<Uint8Array>;

const DEFAULT_MAX_IMAGE_WIDTH_PIXELS = 6.5 * DOCX_DPI; // A4 effective width in pixels

/**
 * DOCX-supported image types (aligned with docx-plus IImageOptions)
 */
export type DocxImageType = "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf" | "svg";

/**
 * Mapping from MIME types / file extensions to DOCX image type strings.
 * Covers all types supported by docx-plus 0.1.2.
 */
const EXTENSION_TO_DOCX_TYPE: Record<string, DocxImageType> = {
  jpg: "jpg",
  jpeg: "jpg",
  png: "png",
  gif: "gif",
  bmp: "bmp",
  tif: "tif",
  tiff: "tif",
  ico: "ico",
  emf: "emf",
  wmf: "wmf",
  svg: "svg",
};

const MIME_TO_DOCX_TYPE: Record<string, DocxImageType> = {
  "image/jpeg": "jpg",
  "image/png": "png",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tif",
  "image/x-icon": "ico",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
  "image/svg+xml": "svg",
};

/**
 * Extract image type from URL or base64 data
 */
export function getImageTypeFromSrc(src: string): DocxImageType {
  if (src.startsWith("data:")) {
    const match = src.match(/data:image\/([\w+-]+);/);
    if (match) {
      return MIME_TO_DOCX_TYPE[match[1].toLowerCase()] || "png";
    }
  } else {
    const extension = src.split(".").pop()?.toLowerCase();
    if (extension) {
      return EXTENSION_TO_DOCX_TYPE[extension] || "png";
    }
  }

  return "png";
}

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
 * Get image width with priority: node attrs > image meta > calculated > default
 *
 * Note: maxWidth constraint only applies to inline (non-floating) images.
 * Floating images maintain their original dimensions.
 */
export function getImageWidth(
  node: { attrs?: { width?: number | null; floating?: any } },
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number {
  if (node.attrs?.width !== undefined && node.attrs?.width !== null) {
    const requestedWidth = node.attrs.width;

    // Only constrain width for inline (non-floating) images
    if (!node.attrs.floating && maxWidth) {
      const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;
      if (maxWidthPixels && requestedWidth > maxWidthPixels) {
        return maxWidthPixels;
      }
    }

    return requestedWidth;
  }

  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.width;
  }

  return maxWidthPixels || DEFAULT_MAX_IMAGE_WIDTH_PIXELS;
}

/**
 * Get image height with priority: node attrs > image meta > calculated > default
 *
 * Note: maxWidth constraint only applies to inline (non-floating) images.
 * Floating images maintain their original dimensions and aspect ratio.
 */
export function getImageHeight(
  node: { attrs?: { height?: number | null; width?: number | null; floating?: any } },
  width: number,
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number {
  if (node.attrs?.height !== undefined && node.attrs?.height !== null) {
    const requestedHeight = node.attrs.height;

    // Only constrain height for inline (non-floating) images when width was also constrained
    if (!node.attrs.floating && maxWidth && node.attrs?.width) {
      const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;
      const requestedWidth = node.attrs.width;

      if (maxWidthPixels && requestedWidth > maxWidthPixels) {
        // Maintain aspect ratio when width is constrained
        const scaleFactor = maxWidthPixels / requestedWidth;
        return Math.round(requestedHeight * scaleFactor);
      }
    }

    return requestedHeight;
  }

  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.height;
  }

  return Math.round(width * 0.75);
}

/**
 * Fetch image data and metadata from HTTP/HTTPS URL
 * (Only for use without custom handler)
 */
export async function getImageDataAndMeta(
  url: string,
): Promise<{ data: Uint8Array; meta: ImageMeta }> {
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
}
