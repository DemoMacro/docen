/**
 * Color conversion utilities for DOCX processing
 * Handles conversion between color names and hex values
 */

/**
 * Color name to hex value mapping
 * Includes common HTML/CSS color names
 */
export const COLOR_NAME_TO_HEX: Record<string, string> = {
  // Basic colors
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",

  // Extended colors
  orange: "#FFA500",
  purple: "#800080",
  pink: "#FFC0CB",
  brown: "#A52A2A",
  gray: "#808080",
  grey: "#808080",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  lime: "#00FF00",
  navy: "#000080",
  teal: "#008080",
  maroon: "#800000",
  olive: "#808000",
  silver: "#C0C0C0",
  gold: "#FFD700",
  indigo: "#4B0082",
  violet: "#EE82EE",

  // Additional common colors
  aqua: "#00FFFF",
  fuchsia: "#FF00FF",
  darkblue: "#00008B",
  darkcyan: "#008B8B",
  darkgrey: "#A9A9A9",
  darkgreen: "#006400",
  darkkhaki: "#BDB76B",
  darkmagenta: "#8B008B",
  darkolivegreen: "#556B2F",
  darkorange: "#FF8C00",
  darkorchid: "#9932CC",
  darkred: "#8B0000",
  darksalmon: "#E9967A",
  darkviolet: "#9400D3",
  lightblue: "#ADD8E6",
  lightcyan: "#E0FFFF",
  lightgreen: "#90EE90",
  lightgrey: "#D3D3D3",
  lightpink: "#FFB6C1",
  lightyellow: "#FFFFE0",
};

/**
 * Convert color name or hex to normalized hex value
 * @param color - Color as name (e.g., "red") or hex (e.g., "#FF0000" or "FF0000")
 * @returns Normalized hex color string (e.g., "#FF0000") or undefined if invalid
 *
 * @example
 * convertColorToHex("red") // returns "#FF0000"
 * convertColorToHex("#FF0000") // returns "#FF0000"
 * convertColorToHex("FF0000") // returns "#FF0000"
 * convertColorToHex("invalid") // returns undefined
 */
export function convertColorToHex(color?: string): string | undefined {
  if (!color) return undefined;

  // Already in hex format
  if (color.startsWith("#")) {
    return color;
  }

  // Try to find in color name mapping (case-insensitive)
  return COLOR_NAME_TO_HEX[color.toLowerCase()] || color;
}

/**
 * Normalize hex color to ensure it has a # prefix
 * @param color - Color value with or without # prefix
 * @returns Hex color with # prefix
 *
 * @example
 * normalizeHexColor("FF0000") // returns "#FF0000"
 * normalizeHexColor("#FF0000") // returns "#FF0000"
 */
export function normalizeHexColor(color: string): string {
  if (color.startsWith("#")) {
    return color;
  }
  return `#${color}`;
}
