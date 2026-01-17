import type { Root, Element } from "xast";
import { fromXml } from "xast-util-from-xml";
import { findChild, findDeepChildren } from "../utils/xml";

/**
 * Character format information from a style definition
 */
export interface CharFormat {
  color?: string; // Hex color with # (e.g., "#FF0000")
  bold?: boolean;
  italic?: boolean;
  fontSize?: number; // Half-points (DOCX unit)
  fontFamily?: string;
  underline?: boolean;
  strike?: boolean;
}

/**
 * Style information from styles.xml
 */
export interface StyleInfo {
  styleId: string;
  name?: string;
  outlineLvl?: number; // 0-9, where 0 is Heading 1, 1 is Heading 2, etc.
  charFormat?: CharFormat; // Character format from style definition
}

export type StyleMap = Map<string, StyleInfo>;

/**
 * Parse styles.xml to build style map
 * Extracts outlineLvl from paragraph styles to identify headings
 * Extracts character format (color, bold, etc.) from style definitions
 */
export function parseStylesXml(files: Record<string, Uint8Array>): StyleMap {
  const styleMap = new Map<string, StyleInfo>();
  const stylesXml = files["word/styles.xml"];
  if (!stylesXml) return styleMap;

  const stylesXast = fromXml(new TextDecoder().decode(stylesXml));
  const styles = findChild(stylesXast, "w:styles");
  if (!styles) return styleMap;

  // Find all paragraph styles
  const paragraphStyles = findDeepChildren(styles, "w:style").filter(
    (style) => style.attributes["w:type"] === "paragraph",
  );

  for (const style of paragraphStyles) {
    const styleId = style.attributes["w:styleId"] as string;
    if (!styleId) continue;

    const styleInfo: StyleInfo = { styleId };

    // Extract style name
    const name = findChild(style, "w:name");
    if (name?.attributes["w:val"]) {
      styleInfo.name = name.attributes["w:val"] as string;
    }

    // Extract outline level (for headings)
    const pPr = findChild(style, "w:pPr");
    if (pPr) {
      const outlineLvl = findChild(pPr, "w:outlineLvl");
      if (outlineLvl?.attributes["w:val"] !== undefined) {
        styleInfo.outlineLvl = parseInt(outlineLvl.attributes["w:val"] as string, 10);
      }
    }

    // Extract character format from style definition
    const rPr = findChild(style, "w:rPr");
    if (rPr) {
      const charFormat: CharFormat = {};

      // Text color
      const color = findChild(rPr, "w:color");
      if (color?.attributes["w:val"] && color.attributes["w:val"] !== "auto") {
        const colorVal = color.attributes["w:val"] as string;
        charFormat.color = colorVal.startsWith("#") ? colorVal : `#${colorVal}`;
      }

      // Bold
      if (findChild(rPr, "w:b")) {
        charFormat.bold = true;
      }

      // Italic
      if (findChild(rPr, "w:i")) {
        charFormat.italic = true;
      }

      // Underline
      if (findChild(rPr, "w:u")) {
        charFormat.underline = true;
      }

      // Strike
      if (findChild(rPr, "w:strike")) {
        charFormat.strike = true;
      }

      // Font size (half-points)
      const sz = findChild(rPr, "w:sz");
      if (sz?.attributes["w:val"]) {
        const sizeVal = sz.attributes["w:val"] as string;
        const size = parseInt(sizeVal, 10);
        if (!isNaN(size)) {
          charFormat.fontSize = size;
        }
      }

      // Font family
      const rFonts = findChild(rPr, "w:rFonts");
      if (rFonts?.attributes["w:ascii"]) {
        charFormat.fontFamily = rFonts.attributes["w:ascii"] as string;
      }

      // Only add charFormat if there's at least one property
      if (Object.keys(charFormat).length > 0) {
        styleInfo.charFormat = charFormat;
      }
    }

    styleMap.set(styleId, styleInfo);
  }

  return styleMap;
}
