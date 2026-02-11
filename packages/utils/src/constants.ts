/**
 * Shared constants for DOCX processing
 * Used across @docen/export-docx and @docen/import-docx packages
 */

/**
 * DOCX DPI (dots per inch) for pixel conversions
 * Word uses 96 DPI internally
 */
export const DOCX_DPI = 96;

/**
 * TWIP (Twentieth of a Point) conversion constants
 * 1 inch = 1440 TWIPs
 */
export const TWIPS_PER_INCH = 1440;

/**
 * EMU (English Metric Unit) conversion constants
 * 1 inch = 914400 EMUs
 */
export const EMUS_PER_INCH = 914400;

/**
 * Font size conversion factors
 * DOCX uses half-points, TipTap uses pixels
 * 1px ≈ 0.75pt, 1pt = 2 half-points
 * So: px * 0.75 * 2 = px * 1.5
 */
export const HALF_POINTS_PER_PIXEL = 1.5;
export const PIXELS_PER_HALF_POINT = 1 / 1.5;

/**
 * Default code font family
 */
export const DEFAULT_CODE_FONT = "Consolas";

/**
 * Checkbox symbols for task lists
 */
export const CHECKBOX_SYMBOLS = {
  checked: "☑",
  unchecked: "☐",
} as const;

/**
 * DOCX style names
 */
export const DOCX_STYLE_NAMES = {
  CODE_BLOCK: "CodeBlock",
  CODE_PREFIX: "Code",
} as const;

/**
 * Text alignment mappings
 */
export const TEXT_ALIGN_MAP = {
  /** TipTap to DOCX alignment mapping */
  tiptapToDocx: {
    left: "left",
    right: "right",
    center: "center",
    justify: "both",
  } as const,
  /** DOCX to TipTap alignment mapping */
  docxToTipTap: {
    left: "left",
    right: "right",
    center: "center",
    both: "justify",
  } as const,
} as const;

/**
 * Common page dimensions in TWIPs
 */
export const PAGE_DIMENSIONS = {
  /** A4 width in TWIPs (8.27 inches = 11906 TWIPs) */
  A4_WIDTH_TWIP: 11906,
  /** Default 1 inch margin in TWIPs */
  DEFAULT_MARGIN_TWIP: 1440,
} as const;
