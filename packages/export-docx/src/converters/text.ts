import { JSONContent } from "@tiptap/core";
import { TextRun, ExternalHyperlink, IRunOptions } from "docx";
import { TextNode } from "@docen/extensions/types";
import { convertColorToHex } from "../utils";

/**
 * Convert TipTap text node to DOCX TextRun or ExternalHyperlink
 */
export function convertText(node: TextNode): TextRun | ExternalHyperlink {
  // Check for marks
  const isBold = node.marks?.some((m) => m.type === "bold");
  const isItalic = node.marks?.some((m) => m.type === "italic");
  const isUnderline = node.marks?.some((m) => m.type === "underline");
  const isStrike = node.marks?.some((m) => m.type === "strike");
  const isCode = node.marks?.some((m) => m.type === "code");
  const isSubscript = node.marks?.some((m) => m.type === "subscript");
  const isSuperscript = node.marks?.some((m) => m.type === "superscript");
  const linkMark = node.marks?.find((m) => m.type === "link");
  const textStyleMark = node.marks?.find((m) => m.type === "textStyle");
  const hasHighlight = node.marks?.some((m) => m.type === "highlight");

  // Handle text color and background color
  const textColor = convertColorToHex(textStyleMark?.attrs?.color);
  const backgroundColor = convertColorToHex(textStyleMark?.attrs?.backgroundColor);

  // Handle font size (convert px to half-points)
  let fontSize: number | undefined;
  if (textStyleMark?.attrs?.fontSize) {
    const fontSizeStr = textStyleMark.attrs.fontSize;
    if (fontSizeStr.endsWith("px")) {
      const px = parseFloat(fontSizeStr);
      if (!isNaN(px)) {
        // Convert px to half-points: 1px ≈ 0.75pt, 1pt = 2 half-points
        // So: px * 0.75 * 2 = px * 1.5
        fontSize = Math.round(px * 1.5);
      }
    }
  }

  // Handle font family (prioritize code font, then textStyle font)
  let fontFamily: string | undefined;
  if (isCode) {
    fontFamily = "Consolas";
  } else if (textStyleMark?.attrs?.fontFamily) {
    fontFamily = textStyleMark.attrs.fontFamily;
  }

  // Build text run options
  const baseOptions: IRunOptions = {
    text: node.text || "",
    bold: isBold || undefined,
    italics: isItalic || undefined,
    underline: isUnderline ? {} : undefined,
    strike: isStrike || undefined,
    font: fontFamily,
    size: fontSize,
    subScript: isSubscript || undefined,
    superScript: isSuperscript || undefined,
    color: textColor,
    shading: backgroundColor ? { fill: backgroundColor } : undefined,
    highlight: hasHighlight ? "yellow" : undefined,
  };

  // Return hyperlink if link mark exists
  if (linkMark?.attrs?.href) {
    return new ExternalHyperlink({
      children: [
        new TextRun({
          ...baseOptions,
          style: "Hyperlink",
        }),
      ],
      link: linkMark.attrs.href,
    });
  }

  // Return regular text run
  return new TextRun(baseOptions);
}

/**
 * Convert TipTap hardBreak node to DOCX TextRun with break
 */
export function convertHardBreak(
  marks?: Array<{ type: string; attrs?: Record<string, any> }>,
): TextRun {
  // Build options object with proper types
  const options: {
    text: string;
    break: number;
    bold?: boolean;
    italics?: boolean;
    underline?: {};
    strike?: boolean;
    color?: string;
  } = {
    text: "",
    break: 1,
  };

  // Apply formatting marks to hardBreak
  if (marks) {
    for (const mark of marks) {
      switch (mark.type) {
        case "bold":
          options.bold = true;
          break;
        case "italic":
          options.italics = true;
          break;
        case "underline":
          options.underline = {}; // Empty object for single underline
          break;
        case "strike":
          options.strike = true;
          break;
        case "textStyle":
          if (mark.attrs?.color) {
            options.color = mark.attrs.color;
          }
          break;
      }
    }
  }

  return new TextRun(options);
}

/**
 * Convert array of text nodes (text, hardBreak) to DOCX elements
 * Returns flattened array of TextRun or ExternalHyperlink
 */
export function convertTextNodes(
  nodes: JSONContent[] = [],
): Array<TextRun | ExternalHyperlink | undefined> {
  return nodes.flatMap((contentNode) => {
    if (contentNode.type === "text") {
      return [convertText(contentNode as TextNode)];
    } else if (contentNode.type === "hardBreak") {
      return [convertHardBreak(contentNode.marks)];
    }
    return [];
  });
}
