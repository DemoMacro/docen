/**
 * Shared utility functions for the export-docx package
 */

import { imageMeta as getImageMetadata, type ImageMeta } from "image-meta";
import { ofetch } from "ofetch";
import type { PositiveUniversalMeasure } from "docx";
import type { DocxExportOptions } from "./option";

/**
 * Constants for image size calculation
 *
 * Note: docx.js uses 96 DPI internally for pixel to EMU conversion
 * EMU conversion: pixels * 9525 = EMU (where 9525 = 914400 / 96)
 */
export const DOCX_DPI = 96; // docx.js internal DPI for pixel to EMU conversion
const DEFAULT_MAX_IMAGE_WIDTH_PIXELS = 6.5 * 96; // A4 effective width in pixels (6.5" * 96 DPI)

/**
 * Convert TWIPs to pixels
 *
 * @param twip - Value in TWIPs
 * @returns Value in pixels
 */
export const convertTwipToPixels = (twip: number): number => {
  return Math.round((twip * DOCX_DPI) / 1440);
};

/**
 * Parse CSS length value to pixels
 * Supports: px, pt, em, rem, %, and unitless values
 * Used for converting TipTap attrs (CSS strings) to pixels for DOCX conversion
 *
 * @param value - CSS length value (e.g., "20px", "1.5rem", "2em", "12pt")
 * @returns Value in pixels
 */
export const convertCssLengthToPixels = (value: string): number => {
  if (!value) return 0;

  // Remove whitespace
  value = value.trim();

  // Match number and optional unit
  const match = value.match(/^([\d.]+)(px|pt|em|rem|%|)?$/);
  if (!match) return 0;

  const num = parseFloat(match[1]);
  if (isNaN(num)) return 0;

  const unit = match[2] || "px";

  // Unit conversion factors to pixels
  const UNIT_TO_PIXELS: Record<string, number> = {
    px: 1,
    pt: 1.333, // 1pt = 1.333px (96/72)
    em: 16, // Assume 16px base font size
    rem: 16, // Same as em
    "%": 0.16, // % of em, assume 16px base (16/100)
  };

  const factor = UNIT_TO_PIXELS[unit] ?? 1;
  return Math.round(num * factor);
};

/**
 * Convert pixels to TWIPs (Twentieth of a Point)
 *
 * @param px - Value in pixels
 * @returns Value in TWIPs
 */
export const convertPixelsToTwip = (px: number): number => {
  return Math.round(px * 15);
};

/**
 * Unit conversion factors to inches
 */
const UNIT_TO_INCHES: Record<string, number> = {
  in: 1,
  mm: 1 / 25.4,
  cm: 1 / 2.54,
  pt: 1 / 72,
  pc: 1 / 6,
  pi: 1 / 6,
} as const;

/**
 * Convert universal measure to inches
 * Compatible with docx.js UniversalMeasure type
 *
 * @param value - Number (inches) or string like "6in", "152.4mm", "15.24cm"
 * @returns Value in inches
 */
export const convertMeasureToInches = (value: number | PositiveUniversalMeasure): number => {
  if (typeof value === "number") {
    // Numbers are treated as inches
    return value;
  }

  // Parse strings like "6in", "152.4mm", "15.24cm"
  const match = value.match(/^([\d.]+)(in|mm|cm|pt|pc|pi)$/);
  if (match) {
    const numValue = parseFloat(match[1]);
    const unit = match[2];

    const factor = UNIT_TO_INCHES[unit];
    return factor !== undefined ? numValue * factor : numValue;
  }

  // Fallback: try to parse as number
  const num = parseFloat(value);
  return isNaN(num) ? 6.5 : num; // Default to 6.5 inches
};

/**
 * Convert universal measure to pixels
 * Compatible with docx.js UniversalMeasure type
 *
 * @param value - Number (pixels) or string like "6in", "152.4mm", "15.24cm"
 * @returns Value in pixels
 */
export const convertMeasureToPixels = (value: number | PositiveUniversalMeasure): number => {
  if (typeof value === "number") {
    // Numbers are treated as pixels for image width
    return value;
  }

  // Reuse convertMeasureToInches and convert to pixels
  const inches = convertMeasureToInches(value);
  return Math.round(inches * DOCX_DPI);
};

/**
 * Calculate effective content width from document options
 *
 * @param options - Export options containing section settings
 * @returns Effective content width in pixels
 */
export function calculateEffectiveContentWidth(options?: DocxExportOptions): number {
  const DEFAULT_PAGE_WIDTH_TWIP = 11905; // A4 width in TWIPs (210mm)
  const DEFAULT_MARGIN_TWIP = 1133; // 20mm margin in TWIPs

  if (!options?.sections || options.sections.length === 0) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const firstSection = options.sections[0];
  if (!firstSection.properties?.page) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const pageSettings = firstSection.properties.page;

  // Get page width (in TWIPs from docx.js)
  let pageWidth = DEFAULT_PAGE_WIDTH_TWIP;
  if (pageSettings.size?.width) {
    const widthValue = pageSettings.size.width;
    pageWidth =
      typeof widthValue === "number"
        ? widthValue
        : Math.round(convertMeasureToInches(widthValue) * 1440);
  }

  // Get margins (in TWIPs from docx.js)
  const marginSettings = pageSettings.margin;
  const marginLeft = marginSettings?.left
    ? typeof marginSettings.left === "number"
      ? marginSettings.left
      : Math.round(convertMeasureToInches(marginSettings.left) * 1440)
    : DEFAULT_MARGIN_TWIP;
  const marginRight = marginSettings?.right
    ? typeof marginSettings.right === "number"
      ? marginSettings.right
      : Math.round(convertMeasureToInches(marginSettings.right) * 1440)
    : DEFAULT_MARGIN_TWIP;

  const effectiveWidth = pageWidth - marginLeft - marginRight;
  return Math.max(convertTwipToPixels(effectiveWidth), DOCX_DPI);
}

/**
 * Calculate appropriate display size for image (mimicking Word's behavior)
 *
 * @param imageMeta - Image metadata
 * @param maxWidthPixels - Maximum available width in pixels
 * @returns Display size in pixels
 */
export function calculateDisplaySize(
  imageMeta: { width?: number; height?: number },
  maxWidthPixels: number = DEFAULT_MAX_IMAGE_WIDTH_PIXELS,
): { width: number; height: number } {
  if (!imageMeta.width || !imageMeta.height) {
    return {
      width: maxWidthPixels,
      height: Math.round(maxWidthPixels * 0.75),
    };
  }

  // If image fits within max width, use original size
  if (imageMeta.width <= maxWidthPixels) {
    return {
      width: imageMeta.width,
      height: imageMeta.height,
    };
  }

  // Scale down proportionally
  const scaleFactor = maxWidthPixels / imageMeta.width;
  return {
    width: maxWidthPixels,
    height: Math.round(imageMeta.height * scaleFactor),
  };
}

/**
 * Extract image type from URL or base64 data
 */
export function getImageTypeFromSrc(src: string): "png" | "jpeg" | "gif" | "bmp" | "tiff" {
  // MIME type to image type mapping
  const MIME_TO_TYPE: Record<string, "png" | "jpeg" | "gif" | "bmp" | "tiff"> = {
    jpg: "jpeg",
    jpeg: "jpeg",
    png: "png",
    gif: "gif",
    bmp: "bmp",
    tiff: "tiff",
  };

  // File extension to image type mapping
  const EXT_TO_TYPE: Record<string, "png" | "jpeg" | "gif" | "bmp" | "tiff"> = {
    jpg: "jpeg",
    jpeg: "jpeg",
    png: "png",
    gif: "gif",
    bmp: "bmp",
    tiff: "tiff",
  };

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
 * Get image width with priority: node attrs > image meta > calculated > default
 *
 * @param node - Image node
 * @param imageMeta - Image metadata
 * @param maxWidth - Maximum available width (number = pixels, or string like "6in", "152.4mm")
 * @returns Image width in pixels
 */
export function getImageWidth(
  node: { attrs?: { width?: number | null } },
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number {
  // Explicit width attribute has highest priority
  if (node.attrs?.width !== undefined && node.attrs?.width !== null) {
    return node.attrs.width;
  }

  // Convert maxWidth to pixels if provided
  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  // Calculate based on metadata and available width
  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.width;
  }

  // Fallback to available width or default
  return maxWidthPixels || DEFAULT_MAX_IMAGE_WIDTH_PIXELS;
}

/**
 * Get image height with priority: node attrs > image meta > calculated > default
 *
 * @param node - Image node
 * @param width - Calculated image width in pixels
 * @param imageMeta - Image metadata
 * @param maxWidth - Maximum available width (number = pixels, or string like "6in", "152.4mm")
 * @returns Image height in pixels
 */
export function getImageHeight(
  node: { attrs?: { height?: number | null } },
  width: number,
  imageMeta?: { width?: number; height?: number },
  maxWidth?: number | PositiveUniversalMeasure,
): number {
  // Explicit height attribute has highest priority
  if (node.attrs?.height !== undefined && node.attrs?.height !== null) {
    return node.attrs.height;
  }

  // Convert maxWidth to pixels if provided
  const maxWidthPixels = maxWidth !== undefined ? convertMeasureToPixels(maxWidth) : undefined;

  // Calculate based on metadata and available width (maintains aspect ratio)
  if (imageMeta?.width && imageMeta?.height) {
    const displaySize = calculateDisplaySize(imageMeta, maxWidthPixels);
    return displaySize.height;
  }

  // Fallback to aspect ratio based on width
  return Math.round(width * 0.75);
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
