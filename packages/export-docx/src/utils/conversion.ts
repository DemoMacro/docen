import type { PositiveUniversalMeasure } from "docx";
import type { DocxExportOptions } from "../options";
import type { Border, Shading } from "@docen/extensions/types";
import {
  convertMeasureToInches,
  convertTwipToPixels,
  DOCX_DPI,
  TWIPS_PER_INCH,
  PAGE_DIMENSIONS,
} from "@docen/utils";

/**
 * Normalize margin value to TWIPs
 * Converts number (already TWIPs) or PositiveUniversalMeasure to TWIPs
 */
const normalizeMarginToTwip = (
  margin: number | PositiveUniversalMeasure | undefined,
  fallback: number,
): number => {
  if (!margin) return fallback;
  return typeof margin === "number"
    ? margin
    : Math.round(convertMeasureToInches(margin) * TWIPS_PER_INCH);
};

/**
 * Calculate effective content width from document options
 */
export function calculateEffectiveContentWidth(options?: DocxExportOptions): number {
  const DEFAULT_PAGE_WIDTH_TWIP = PAGE_DIMENSIONS.A4_WIDTH_TWIP;
  const DEFAULT_MARGIN_TWIP = PAGE_DIMENSIONS.DEFAULT_MARGIN_TWIP;

  if (!options?.sections?.length) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const firstSection = options.sections[0];
  if (!firstSection.properties?.page) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const pageSettings = firstSection.properties.page;

  let pageWidth: number = DEFAULT_PAGE_WIDTH_TWIP;
  if (pageSettings.size?.width) {
    const widthValue = pageSettings.size.width;
    pageWidth =
      typeof widthValue === "number"
        ? widthValue
        : Math.round(convertMeasureToInches(widthValue) * TWIPS_PER_INCH);
  }

  const marginSettings = pageSettings.margin;
  const marginLeft = normalizeMarginToTwip(marginSettings?.left, DEFAULT_MARGIN_TWIP);
  const marginRight = normalizeMarginToTwip(marginSettings?.right, DEFAULT_MARGIN_TWIP);

  const effectiveWidth = pageWidth - marginLeft - marginRight;
  return Math.max(convertTwipToPixels(effectiveWidth), DOCX_DPI);
}

/**
 * Convert Border to docx.js format
 */
export function convertBorder(border?: Border):
  | {
      color?: string;
      size?: number;
      style?: string;
      space?: number;
    }
  | undefined {
  if (!border) return undefined;

  const docxBorder: {
    color?: string;
    size?: number;
    style?: string;
    space?: number;
  } = {};

  if (border.color) {
    // Remove # prefix for docx.js
    docxBorder.color = border.color.replace("#", "");
  }

  if (border.size !== undefined) {
    // Keep as eighth-points (DOCX native unit)
    docxBorder.size = border.size;
  }

  if (border.style) {
    docxBorder.style = border.style;
  }

  if (border.space !== undefined) {
    docxBorder.space = border.space;
  }

  return Object.keys(docxBorder).length > 0 ? docxBorder : undefined;
}

/**
 * Convert Shading to docx.js format
 */
export function convertShading(shading?: Shading):
  | {
      fill?: string;
      color?: string;
      val?: string;
    }
  | undefined {
  if (!shading || !shading.fill) return undefined;

  const docxShading: {
    fill?: string;
    color?: string;
    val?: string;
  } = {};

  if (shading.fill) {
    // Remove # prefix for docx.js
    docxShading.fill = shading.fill.replace("#", "");
  }

  if (shading.color) {
    // Remove # prefix for docx.js
    docxShading.color = shading.color.replace("#", "");
  }

  // Default to "clear" for solid fill
  docxShading.val = shading.type || "clear";

  return docxShading;
}
