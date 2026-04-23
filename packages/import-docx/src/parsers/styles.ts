import { fromXml } from "xast-util-from-xml";
import { findChild, findDeepChildren, parseTwipAttr, convertTwipToCssString } from "@docen/utils";
import type { Element } from "xast";
import type { Border, Shading } from "@docen/extensions/types";

/**
 * Character format information from a style definition
 */
export interface CharFormat {
  color?: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFamily?: string;
  underline?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  characterSpacing?: number;
  rtl?: boolean;
}

/**
 * Paragraph format information from a style definition
 */
export interface ParagraphFormat {
  shading?: Shading;
  borderTop?: Border;
  borderBottom?: Border;
  borderLeft?: Border;
  borderRight?: Border;
}

/**
 * Style information from styles.xml
 */
export interface StyleInfo {
  styleId: string;
  name?: string;
  outlineLvl?: number; // 0-9, where 0 is Heading 1, 1 is Heading 2, etc.
  charFormat?: CharFormat; // Character format from style definition
  paragraphFormat?: ParagraphFormat; // Paragraph format from style definition
}

export type StyleMap = Map<string, StyleInfo>;

/**
 * Parse a single border element
 */
export function parseBorder(borderNode: Element | null): Border | null {
  if (!borderNode) return null;

  const val = borderNode.attributes["w:val"] as string;
  const size = borderNode.attributes["w:sz"] as string;
  const color = borderNode.attributes["w:color"] as string;
  const space = borderNode.attributes["w:space"] as string;

  // Map DOCX border styles
  const styleMap: Record<string, Border["style"]> = {
    single: "single",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    dotDash: "dotDash",
    dotDotDash: "dotDotDash",
    none: "none",
    nil: "none",
  };

  const border: Border = {};

  if (color && color !== "auto") {
    border.color = `#${color}`;
  }

  if (size) {
    // Keep as eighth-points (DOCX native unit)
    border.size = parseInt(size);
  }

  if (val && styleMap[val]) {
    border.style = styleMap[val];
  }

  if (space) {
    border.space = parseInt(space);
  }

  return Object.keys(border).length > 0 ? border : null;
}

/**
 * Parse borders from w:pBdr element
 */
export function parseBorders(pPr: Element | null): {
  borderTop?: Border;
  borderBottom?: Border;
  borderLeft?: Border;
  borderRight?: Border;
} | null {
  if (!pPr) return null;

  const borderElement = findChild(pPr, "w:pBdr");
  if (!borderElement) return null;

  const borders: {
    borderTop?: Border;
    borderBottom?: Border;
    borderLeft?: Border;
    borderRight?: Border;
  } = {};

  const topBorder = parseBorder(findChild(borderElement, "w:top"));
  if (topBorder) borders.borderTop = topBorder;

  const bottomBorder = parseBorder(findChild(borderElement, "w:bottom"));
  if (bottomBorder) borders.borderBottom = bottomBorder;

  const leftBorder = parseBorder(findChild(borderElement, "w:left"));
  if (leftBorder) borders.borderLeft = leftBorder;

  const rightBorder = parseBorder(findChild(borderElement, "w:right"));
  if (rightBorder) borders.borderRight = rightBorder;

  return Object.keys(borders).length > 0 ? borders : null;
}

/**
 * Parse shading from w:shd element
 */
export function parseShading(pPr: Element | null): Shading | null {
  if (!pPr) return null;

  const shd = findChild(pPr, "w:shd");
  if (!shd) return null;

  const shading: Shading = {};

  if (shd.attributes["w:fill"]) {
    const fill = shd.attributes["w:fill"] as string;
    shading.fill = fill.startsWith("#") ? fill : `#${fill}`;
  }

  if (shd.attributes["w:color"] && shd.attributes["w:color"] !== "auto") {
    const color = shd.attributes["w:color"] as string;
    shading.color = color.startsWith("#") ? color : `#${color}`;
  }

  if (shd.attributes["w:val"]) {
    shading.type = shd.attributes["w:val"] as string;
  }

  return Object.keys(shading).length > 0 ? shading : null;
}

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

    // Extract outline level (for headings) and paragraph format
    const pPr = findChild(style, "w:pPr");
    if (pPr) {
      const outlineLvl = findChild(pPr, "w:outlineLvl");
      if (outlineLvl?.attributes["w:val"] !== undefined) {
        styleInfo.outlineLvl = parseInt(outlineLvl.attributes["w:val"] as string, 10);
      }

      // Extract paragraph format (borders, shading)
      const borders = parseBorders(pPr);
      const shading = parseShading(pPr);

      if (borders || shading) {
        const paragraphFormat: ParagraphFormat = {};
        if (borders) Object.assign(paragraphFormat, borders);
        if (shading) paragraphFormat.shading = shading;

        // Only add if there's at least one property
        if (Object.keys(paragraphFormat).length > 0) {
          styleInfo.paragraphFormat = paragraphFormat;
        }
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
      const bold = findChild(rPr, "w:b");
      if (bold) {
        const val = bold.attributes["w:val"];
        if (val !== "0" && val !== "false") {
          charFormat.bold = true;
        }
      }

      // Italic
      const italic = findChild(rPr, "w:i");
      if (italic) {
        const val = italic.attributes["w:val"];
        if (val !== "0" && val !== "false") {
          charFormat.italic = true;
        }
      }

      // Underline
      const underline = findChild(rPr, "w:u");
      if (underline) {
        const val = underline.attributes["w:val"];
        if (val !== "none" && val !== "false" && val !== "0") {
          charFormat.underline = true;
        }
      }

      // Strike
      const strike = findChild(rPr, "w:strike");
      if (strike) {
        const val = strike.attributes["w:val"];
        if (val !== "0" && val !== "false") {
          charFormat.strike = true;
        }
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

      // Double strikethrough
      const dstrike = findChild(rPr, "w:dstrike");
      if (dstrike) {
        const val = dstrike.attributes["w:val"];
        if (val !== "0" && val !== "false") {
          charFormat.doubleStrike = true;
        }
      }

      // Character spacing (w:spacing in rPr, unit: twips)
      const spacing = findChild(rPr, "w:spacing");
      if (spacing?.attributes["w:val"]) {
        const val = parseInt(spacing.attributes["w:val"] as string);
        if (!isNaN(val)) {
          charFormat.characterSpacing = val;
        }
      }

      // Right-to-left text
      const rtl = findChild(rPr, "w:rtl");
      if (rtl) {
        const val = rtl.attributes["w:val"];
        if (val !== "0" && val !== "false") {
          charFormat.rtl = true;
        }
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

/**
 * Extract all paragraph style attributes from a paragraph element
 * Merges direct paragraph properties with style-based properties
 */
export function extractParagraphStyles(
  node: Element,
  styleInfo?: StyleInfo,
): {
  indentLeft?: string;
  indentRight?: string;
  indentFirstLine?: string;
  spacingBefore?: string;
  spacingAfter?: string;
  shading?: Shading;
  borderTop?: Border;
  borderBottom?: Border;
  borderLeft?: Border;
  borderRight?: Border;
} | null {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return null;

  const result: Record<string, unknown> = {};

  // Start with style-based properties (if available)
  if (styleInfo?.paragraphFormat) {
    const pf = styleInfo.paragraphFormat;
    if (pf.shading) result.shading = pf.shading;
    if (pf.borderTop) result.borderTop = pf.borderTop;
    if (pf.borderBottom) result.borderBottom = pf.borderBottom;
    if (pf.borderLeft) result.borderLeft = pf.borderLeft;
    if (pf.borderRight) result.borderRight = pf.borderRight;
  }

  // Extract indentation
  const ind = findChild(pPr, "w:ind");
  if (ind) {
    const left =
      parseTwipAttr(ind.attributes, "w:left") || parseTwipAttr(ind.attributes, "w:start");
    if (left) {
      const leftTwip = parseInt(left, 10);
      result.indentLeft = convertTwipToCssString(leftTwip);
    }

    const right =
      parseTwipAttr(ind.attributes, "w:right") || parseTwipAttr(ind.attributes, "w:end");
    if (right) {
      const rightTwip = parseInt(right, 10);
      result.indentRight = convertTwipToCssString(rightTwip);
    }

    const firstLine = parseTwipAttr(ind.attributes, "w:firstLine");
    if (firstLine) {
      const firstLineTwip = parseInt(firstLine, 10);
      result.indentFirstLine = convertTwipToCssString(firstLineTwip);
    } else {
      const hanging = parseTwipAttr(ind.attributes, "w:hanging");
      if (hanging) {
        const leftTwip = left ? parseInt(left, 10) : 0;
        const hangingTwip = parseInt(hanging, 10);
        const firstLineTwip = leftTwip - hangingTwip;
        result.indentFirstLine = convertTwipToCssString(firstLineTwip);
      }
    }
  }

  // Extract spacing
  const spacing = findChild(pPr, "w:spacing");
  if (spacing) {
    const before = parseTwipAttr(spacing.attributes, "w:before");
    if (before) {
      const beforeTwip = parseInt(before, 10);
      result.spacingBefore = convertTwipToCssString(beforeTwip);
    }

    const after = parseTwipAttr(spacing.attributes, "w:after");
    if (after) {
      const afterTwip = parseInt(after, 10);
      result.spacingAfter = convertTwipToCssString(afterTwip);
    }
  }

  // Extract shading (direct properties override style-based)
  const shading = parseShading(pPr);
  if (shading) {
    result.shading = shading;
  }

  // Extract borders (direct properties override style-based)
  const borders = parseBorders(pPr);
  if (borders) {
    Object.assign(result, borders);
  }

  return Object.keys(result).length > 0
    ? (result as ReturnType<typeof extractParagraphStyles>)
    : null;
}
