import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "../option";
import { extractRuns, extractAlignment } from "./text";
import { findChild } from "../utils/xml";

/**
 * Convert TWIPs to pixels
 * 1 inch = 1440 TWIPs, 1px ≈ 15 TWIPs (at 96 DPI: 1px = 0.75pt = 15 TWIP)
 */
function twipToPx(twip: number): number {
  return Math.round(twip / 15);
}

/**
 * Extract paragraph style attributes from DOCX paragraph properties
 */
function extractParagraphStyles(node: Element): {
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  spacingBefore?: number;
  spacingAfter?: number;
} | null {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return null;

  const result: {
    indentLeft?: number;
    indentRight?: number;
    indentFirstLine?: number;
    spacingBefore?: number;
    spacingAfter?: number;
  } = {};

  // Extract indentation from w:ind
  const ind = findChild(pPr, "w:ind");
  if (ind) {
    const left = ind.attributes["w:left"];
    const right = ind.attributes["w:right"];
    const firstLine = ind.attributes["w:firstLine"];
    const hanging = ind.attributes["w:hanging"];

    if (typeof left === "string") {
      const leftValue = parseInt(left, 10);
      if (!isNaN(leftValue)) result.indentLeft = twipToPx(leftValue);
    }

    if (typeof right === "string") {
      const rightValue = parseInt(right, 10);
      if (!isNaN(rightValue)) result.indentRight = twipToPx(rightValue);
    }

    if (typeof firstLine === "string") {
      const firstLineValue = parseInt(firstLine, 10);
      if (!isNaN(firstLineValue)) result.indentFirstLine = twipToPx(firstLineValue);
    } else if (typeof hanging === "string") {
      // Convert hanging indent to negative first line indent
      const hangingValue = parseInt(hanging, 10);
      if (!isNaN(hangingValue)) result.indentFirstLine = -twipToPx(hangingValue);
    }
  }

  // Extract spacing from w:spacing
  const spacing = findChild(pPr, "w:spacing");
  if (spacing) {
    const before = spacing.attributes["w:before"];
    const after = spacing.attributes["w:after"];

    if (typeof before === "string") {
      const beforeValue = parseInt(before, 10);
      if (!isNaN(beforeValue)) result.spacingBefore = twipToPx(beforeValue);
    }

    if (typeof after === "string") {
      const afterValue = parseInt(after, 10);
      if (!isNaN(afterValue)) result.spacingAfter = twipToPx(afterValue);
    }
  }

  // Return null if no styles found
  return Object.keys(result).length > 0 ? result : null;
}

/**
 * Convert DOCX paragraph node to TipTap paragraph
 */
export async function convertParagraph(
  node: Element,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
): Promise<JSONContent> {
  // Check if it's a heading by finding w:pPr > w:pStyle
  const pPr = findChild(node, "w:pPr");
  let styleName: string | undefined;
  if (pPr) {
    const pStyle = findChild(pPr, "w:pStyle");
    if (pStyle) {
      styleName = pStyle.attributes["w:val"] as string;
    }
  }

  if (styleName) {
    // Check if it's a heading (e.g., "Heading1", "Heading2")
    const headingMatch = styleName.match(/^Heading(\d)$/);
    if (headingMatch) {
      const level = parseInt(headingMatch[1]) as 1 | 2 | 3 | 4 | 5 | 6;
      return convertHeading(node, params, level);
    }
  }

  // Extract runs (text, images, hardBreaks)
  const runs = await extractRuns(node, params);

  // Check if this is a horizontal rule (page break)
  if (runs.length === 1 && runs[0].type === "hardBreak") {
    // Check if it's a page break type
    const run = findChild(node, "w:r");
    if (run) {
      const br = findChild(run, "w:br");
      if (br && br.attributes["w:type"] === "page") {
        return { type: "horizontalRule" };
      }
    }
  }

  // Check if paragraph contains only an image (no text)
  // In this case, return the image node directly instead of wrapping in paragraph
  if (runs.length === 1 && runs[0].type === "image") {
    return runs[0] as JSONContent;
  }

  // Regular paragraph
  const attrs = extractAlignment(node);
  const paragraphStyles = extractParagraphStyles(node);

  // Merge alignment and paragraph styles
  const mergedAttrs = {
    ...attrs,
    ...paragraphStyles,
  };

  return {
    type: "paragraph",
    ...(Object.keys(mergedAttrs).length > 0 && { attrs: mergedAttrs }),
    content: runs,
  };
}

/**
 * Convert to heading (internal function)
 */
async function convertHeading(
  node: Element,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
  level: 1 | 2 | 3 | 4 | 5 | 6,
): Promise<JSONContent> {
  const paragraphStyles = extractParagraphStyles(node);

  return {
    type: "heading",
    attrs: {
      level,
      ...paragraphStyles,
    },
    content: await extractRuns(node, params),
  };
}
