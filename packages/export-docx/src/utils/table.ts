import type { TableCellBorder } from "@docen/extensions/types";

/**
 * Convert TipTap border to DOCX border format
 *
 * @param border - TipTap table cell border definition
 * @returns DOCX border options or undefined if no border
 */
export function convertBorder(
  border: TableCellBorder | null | undefined,
):
  | { color?: string; size?: number; style: "single" | "dashed" | "dotted" | "double" | "none" }
  | undefined {
  if (!border) return undefined;

  const styleMap: Record<string, "single" | "dashed" | "dotted" | "double" | "none"> = {
    solid: "single",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    none: "none",
  };

  const docxStyle = border.style ? styleMap[border.style] || "single" : "single";
  const color = border.color?.replace("#", "") || "auto";
  const size = border.width ? border.width * 6 : 4; // Convert pixels to eighth-points

  return {
    color,
    size,
    style: docxStyle,
  };
}
