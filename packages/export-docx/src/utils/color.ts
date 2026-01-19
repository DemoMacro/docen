/**
 * Color name to hex value mapping
 */

const COLOR_NAME_TO_HEX: Record<string, string> = {
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  orange: "#FFA500",
  purple: "#800080",
  pink: "#FFC0CB",
  brown: "#A52A2A",
  black: "#000000",
  white: "#FFFFFF",
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
};

/**
 * Convert color name or hex to normalized hex value
 */
export function convertColorToHex(color?: string): string | undefined {
  if (!color) return undefined;
  if (color.startsWith("#")) return color;
  return COLOR_NAME_TO_HEX[color.toLowerCase()] || color;
}
