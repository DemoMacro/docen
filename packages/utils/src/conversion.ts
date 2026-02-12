/**
 * Unit conversion utilities for DOCX processing
 * Handles conversions between TWIPs, EMUs, pixels, and other units
 */

import { DOCX_DPI, TWIPS_PER_INCH, EMUS_PER_INCH } from "./constants";

const PIXELS_PER_INCH = DOCX_DPI; // 96 DPI

/**
 * Convert TWIPs to CSS pixels (returns number)
 * @param twip - Value in TWIPs (1 inch = 1440 TWIPs)
 * @returns Number value in pixels
 *
 * @example
 * convertTwipToPixels(1440) // returns 96
 */
export function convertTwipToPixels(twip: number): number {
  return Math.round((twip * PIXELS_PER_INCH) / TWIPS_PER_INCH);
}

/**
 * Convert TWIPs to CSS string (returns "px" string)
 * @param twip - Value in TWIPs
 * @returns CSS value string in pixels (e.g., "20px")
 *
 * @example
 * convertTwipToCssString(1440) // returns "96px"
 */
export function convertTwipToCssString(twip: number): string {
  const px = convertTwipToPixels(twip);
  return `${px}px`;
}

/**
 * Convert pixels to TWIPs
 * @param px - Value in pixels
 * @returns Value in TWIPs
 *
 * @example
 * convertPixelsToTwip(96) // returns 1440
 */
export function convertPixelsToTwip(px: number): number {
  return Math.round(px * (TWIPS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert EMUs to pixels
 * EMU = English Metric Unit (1 inch = 914400 EMUs)
 * @param emu - Value in EMUs
 * @returns Value in pixels
 *
 * @example
 * convertEmuToPixels(914400) // returns 96
 */
export function convertEmuToPixels(emu: number): number {
  return Math.round(emu / (EMUS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert pixels to EMUs
 * @param px - Value in pixels
 * @returns Value in EMUs
 *
 * @example
 * convertPixelsToEmu(96) // returns 914400
 */
export function convertPixelsToEmu(px: number): number {
  return Math.round(px * (EMUS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert EMU string to pixels
 * @param emuStr - EMU value as string
 * @returns Pixel value or undefined if invalid
 *
 * @example
 * convertEmuStringToPixels("914400") // returns 96
 * convertEmuStringToPixels("invalid") // returns undefined
 */
export function convertEmuStringToPixels(emuStr: string): number | undefined {
  const emu = parseInt(emuStr, 10);
  if (isNaN(emu)) return undefined;
  return convertEmuToPixels(emu);
}

/**
 * Regular expression for matching CSS length values
 * Supports: px, pt, em, rem, %, and unitless values (including negative values)
 */
const CSS_LENGTH_REGEX = /^(-?[\d.]+)(px|pt|em|rem|%|)?$/;

/**
 * Conversion factors from various units to pixels
 */
const UNIT_TO_PIXELS: Record<string, number> = {
  px: 1,
  pt: 1.333,
  em: 16,
  rem: 16,
  "%": 0.16,
};

/**
 * Parse CSS length value to pixels
 * Supports: px, pt, em, rem, %, and unitless values
 * @param value - CSS length value (e.g., "20px", "1.5em", "100%")
 * @returns Value in pixels
 *
 * @example
 * convertCssLengthToPixels("20px") // returns 20
 * convertCssLengthToPixels("1.5em") // returns 24
 * convertCssLengthToPixels("100%") // returns 16
 * convertCssLengthToPixels("20") // returns 20 (unitless treated as px)
 */
export function convertCssLengthToPixels(value: string): number {
  if (!value) return 0;

  value = value.trim();
  const match = value.match(CSS_LENGTH_REGEX);
  if (!match) return 0;

  const num = parseFloat(match[1]);
  if (isNaN(num)) return 0;

  const unit = match[2] || "px";
  const factor = UNIT_TO_PIXELS[unit] ?? 1;

  return Math.round(num * factor);
}

/**
 * Regular expression for matching universal measure values
 * Used in DOCX for specifying dimensions in various units
 */
const MEASURE_REGEX = /^([\d.]+)(in|mm|cm|pt|pc|pi)$/;

/**
 * Conversion factors from various units to inches
 */
const UNIT_TO_INCHES: Record<string, number> = {
  in: 1,
  mm: 1 / 25.4,
  cm: 1 / 2.54,
  pt: 1 / 72,
  pc: 1 / 6,
  pi: 1 / 6,
};

/**
 * Type for universal measure values in DOCX
 * Can be a number (already in inches) or a string with unit suffix
 */
export type PositiveUniversalMeasure = `${number}${"in" | "mm" | "cm" | "pt" | "pc" | "pi"}`;

/**
 * Convert universal measure to inches
 * Compatible with docx.js UniversalMeasure type
 * @param value - Value in various units (number or string)
 * @returns Value in inches
 *
 * @example
 * convertMeasureToInches(6.5) // returns 6.5
 * convertMeasureToInches("1in") // returns 1
 * convertMeasureToInches("25.4mm") // returns ~1
 */
export function convertMeasureToInches(value: number | PositiveUniversalMeasure): number {
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
}

/**
 * Convert universal measure to pixels
 * Compatible with docx.js UniversalMeasure type
 * @param value - Value in various units (number or string)
 * @returns Value in pixels
 *
 * @example
 * convertMeasureToPixels("1in") // returns 96
 * convertMeasureToPixels(6.5) // returns 624
 */
export function convertMeasureToPixels(value: number | PositiveUniversalMeasure): number {
  if (typeof value === "number") {
    return value;
  }

  const inches = convertMeasureToInches(value);
  return Math.round(inches * DOCX_DPI);
}
