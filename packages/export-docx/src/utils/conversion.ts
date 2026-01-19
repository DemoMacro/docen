import type { PositiveUniversalMeasure } from "docx";
import type { DocxExportOptions } from "../options";

/**
 * Constants for unit conversion
 */

export const DOCX_DPI = 96; // docx.js internal DPI for pixel to EMU conversion

const CSS_LENGTH_REGEX = /^([\d.]+)(px|pt|em|rem|%|)?$/;
const MEASURE_REGEX = /^([\d.]+)(in|mm|cm|pt|pc|pi)$/;

const UNIT_TO_PIXELS: Record<string, number> = {
  px: 1,
  pt: 1.333,
  em: 16,
  rem: 16,
  "%": 0.16,
};

const UNIT_TO_INCHES: Record<string, number> = {
  in: 1,
  mm: 1 / 25.4,
  cm: 1 / 2.54,
  pt: 1 / 72,
  pc: 1 / 6,
  pi: 1 / 6,
};

/**
 * Convert TWIPs to pixels
 */

export const convertTwipToPixels = (twip: number): number => {
  return Math.round((twip * DOCX_DPI) / 1440);
};

/**
 * Parse CSS length value to pixels
 * Supports: px, pt, em, rem, %, and unitless values
 */

export const convertCssLengthToPixels = (value: string): number => {
  if (!value) return 0;

  value = value.trim();
  const match = value.match(CSS_LENGTH_REGEX);
  if (!match) return 0;

  const num = parseFloat(match[1]);
  if (isNaN(num)) return 0;

  const unit = match[2] || "px";
  const factor = UNIT_TO_PIXELS[unit] ?? 1;

  return Math.round(num * factor);
};

/**
 * Convert pixels to TWIPs (Twentieth of a Point)
 */

export const convertPixelsToTwip = (px: number): number => {
  return Math.round(px * 15);
};

/**
 * Convert universal measure to inches
 * Compatible with docx.js UniversalMeasure type
 */

export const convertMeasureToInches = (value: number | PositiveUniversalMeasure): number => {
  if (typeof value === "number") {
    return value;
  }

  const match = value.match(MEASURE_REGEX);
  if (match) {
    const numValue = parseFloat(match[1]);
    const unit = match[2];
    const factor = UNIT_TO_INCHES[unit];

    return factor !== undefined ? numValue * factor : numValue;
  }

  const num = parseFloat(value);
  return isNaN(num) ? 6.5 : num;
};

/**
 * Convert universal measure to pixels
 * Compatible with docx.js UniversalMeasure type
 */

export const convertMeasureToPixels = (value: number | PositiveUniversalMeasure): number => {
  if (typeof value === "number") {
    return value;
  }

  const inches = convertMeasureToInches(value);
  return Math.round(inches * DOCX_DPI);
};

/**
 * Calculate effective content width from document options
 */

export const calculateEffectiveContentWidth = (options?: DocxExportOptions): number => {
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
};
