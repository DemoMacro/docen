import type { PositiveUniversalMeasure } from "docx";
import type { DocxExportOptions } from "../options";
import { convertMeasureToInches, convertTwipToPixels, DOCX_DPI } from "@docen/utils";

/**
 * Normalize margin value to TWIPs
 * Converts number (already TWIPs) or PositiveUniversalMeasure to TWIPs
 */
const normalizeMarginToTwip = (
  margin: number | PositiveUniversalMeasure | undefined,
  fallback: number,
): number => {
  if (!margin) return fallback;
  return typeof margin === "number" ? margin : Math.round(convertMeasureToInches(margin) * 1440);
};

/**
 * Calculate effective content width from document options
 */
export function calculateEffectiveContentWidth(options?: DocxExportOptions): number {
  const DEFAULT_PAGE_WIDTH_TWIP = 11906; // A4 width in TWIPs
  const DEFAULT_MARGIN_TWIP = 1440; // 1 inch margin in TWIPs

  if (!options?.sections?.length) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const firstSection = options.sections[0];
  if (!firstSection.properties?.page) {
    return convertTwipToPixels(DEFAULT_PAGE_WIDTH_TWIP - DEFAULT_MARGIN_TWIP * 2);
  }

  const pageSettings = firstSection.properties.page;

  let pageWidth = DEFAULT_PAGE_WIDTH_TWIP;
  if (pageSettings.size?.width) {
    const widthValue = pageSettings.size.width;
    pageWidth =
      typeof widthValue === "number"
        ? widthValue
        : Math.round(convertMeasureToInches(widthValue) * 1440);
  }

  const marginSettings = pageSettings.margin;
  const marginLeft = normalizeMarginToTwip(marginSettings?.left, DEFAULT_MARGIN_TWIP);
  const marginRight = normalizeMarginToTwip(marginSettings?.right, DEFAULT_MARGIN_TWIP);

  const effectiveWidth = pageWidth - marginLeft - marginRight;
  return Math.max(convertTwipToPixels(effectiveWidth), DOCX_DPI);
}
