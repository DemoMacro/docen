import type { Border } from "./types";

/**
 * Render a Border to CSS string
 */
export function renderBorderCSS(border: Border): string | null {
  if (!border || border.style === "none") return null;
  const size = border.size != null ? `${border.size / 8}pt` : "1pt";
  const color = border.color || "auto";
  const styleMap: Record<string, string> = {
    single: "solid",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    dotDash: "dashed",
    dotDotDash: "dotted",
  };
  const cssStyle = styleMap[border.style || "single"] || "solid";
  return `${cssStyle} ${size} ${color}`;
}
