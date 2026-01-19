/**
 * Unit conversion utilities for DOCX import
 * These convert DOCX internal units to web-friendly formats
 */

/**
 * DOCX DPI (dots per inch) for pixel conversions
 * Word uses 96 DPI internally
 */
export const DOCX_DPI = 96;

/**
 * Conversion factors for DOCX units
 */
const TWIPS_PER_INCH = 1440; // 1 inch = 1440 TWIPs
const EMUS_PER_INCH = 914400; // 1 inch = 914400 EMUs
const PIXELS_PER_INCH = DOCX_DPI; // 96 DPI

/**
 * Convert TWIPs to CSS pixels
 * @param twip - Value in TWIPs (1 inch = 1440 TWIPs)
 * @returns CSS value string in pixels (e.g., "20px")
 */
export function convertTwipToPixels(twip: number): string {
  const px = Math.round((twip * PIXELS_PER_INCH) / TWIPS_PER_INCH);
  return `${px}px`;
}

/**
 * Convert TWIPs to number value in pixels
 * @param twip - Value in TWIPs
 * @returns Number value in pixels
 */
export function convertTwipToPixelNumber(twip: number): number {
  return Math.round((twip * PIXELS_PER_INCH) / TWIPS_PER_INCH);
}

/**
 * Convert pixels to TWIPs
 * @param px - Value in pixels
 * @returns Value in TWIPs
 */
export function convertPixelsToTwip(px: number): number {
  return Math.round(px * (TWIPS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert EMUs to pixels
 * EMU = English Metric Unit (1 inch = 914400 EMUs)
 * @param emu - Value in EMUs
 * @returns Value in pixels
 */
export function convertEmuToPixels(emu: number): number {
  return Math.round(emu / (EMUS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert pixels to EMUs
 * @param px - Value in pixels
 * @returns Value in EMUs
 */
export function convertPixelsToEmu(px: number): number {
  return Math.round(px * (EMUS_PER_INCH / PIXELS_PER_INCH));
}

/**
 * Convert EMU string to pixels
 * @param emuStr - EMU value as string
 * @returns Pixel value or undefined if invalid
 */
export function convertEmuStringToPixels(emuStr: string): number | undefined {
  const emu = parseInt(emuStr, 10);
  if (isNaN(emu)) return undefined;
  return convertEmuToPixels(emu);
}
