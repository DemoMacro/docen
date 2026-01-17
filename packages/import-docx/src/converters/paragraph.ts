import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "../option";
import type { StyleMap, StyleInfo } from "../parsing/styles";
import type { ImageInfo } from "../parsing/types";
import { extractRuns, extractAlignment } from "./text";
import { findChild } from "../utils/xml";

/**
 * Convert TWIPs to CSS pixels
 * 1 inch = 1440 TWIPs, 1px ≈ 15 TWIPs (at 96 DPI: 1px = 0.75pt = 15 TWIP)
 * @param twip - Value in TWIPs
 * @returns CSS value in pixels (e.g., "20px")
 */
function convertTwipToPixels(twip: number): string {
  const px = Math.round(twip / 15);
  return `${px}px`;
}

/**
 * Extract paragraph style attributes from DOCX paragraph properties
 */
function extractParagraphStyles(node: Element): {
  indentLeft?: string;
  indentRight?: string;
  indentFirstLine?: string;
  spacingBefore?: string;
  spacingAfter?: string;
} | null {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return null;

  const result: Record<string, string> = {};

  // Extract indentation from w:ind
  const ind = findChild(pPr, "w:ind");
  if (ind) {
    const parseAttr = (attr: string) => {
      const value = ind.attributes[attr];
      if (typeof value !== "string") return null;
      const num = parseInt(value, 10);
      return isNaN(num) ? null : convertTwipToPixels(num);
    };

    const left = parseAttr("w:left");
    if (left) result.indentLeft = left;

    const right = parseAttr("w:right");
    if (right) result.indentRight = right;

    const firstLine = parseAttr("w:firstLine");
    if (firstLine) {
      result.indentFirstLine = firstLine;
    } else {
      const hanging = parseAttr("w:hanging");
      if (hanging) result.indentFirstLine = `-${hanging}`;
    }
  }

  // Extract spacing from w:spacing
  const spacing = findChild(pPr, "w:spacing");
  if (spacing) {
    const parseAttr = (attr: string) => {
      const value = spacing.attributes[attr];
      if (typeof value !== "string") return null;
      const num = parseInt(value, 10);
      return isNaN(num) ? null : convertTwipToPixels(num);
    };

    const before = parseAttr("w:before");
    if (before) result.spacingBefore = before;

    const after = parseAttr("w:after");
    if (after) result.spacingAfter = after;
  }

  return Object.keys(result).length ? result : null;
}

/**
 * Convert DOCX paragraph node to TipTap paragraph
 */
export async function convertParagraph(
  node: Element,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, ImageInfo>;
    options?: DocxImportOptions;
    styleMap?: StyleMap;
  },
): Promise<JSONContent> {
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const styleName = pStyle?.attributes["w:val"] as string | undefined;

  // Check if it's a heading
  if (styleName && params.styleMap) {
    const styleInfo = params.styleMap.get(styleName);

    // Check outlineLvl (reliable heading indicator)
    if (
      styleInfo?.outlineLvl !== undefined &&
      styleInfo.outlineLvl >= 0 &&
      styleInfo.outlineLvl <= 5
    ) {
      const level = (styleInfo.outlineLvl + 1) as 1 | 2 | 3 | 4 | 5 | 6;
      return convertHeading(node, params, styleInfo, level);
    }

    // Fallback: Check style name pattern
    const headingMatch = styleName.match(/^Heading(\d+)$/);
    if (headingMatch) {
      const level = parseInt(headingMatch[1], 10) as 1 | 2 | 3 | 4 | 5 | 6;
      return convertHeading(node, params, styleInfo, level);
    }
  }

  const styleInfo = styleName && params.styleMap ? params.styleMap.get(styleName) : undefined;
  const runs = await extractRuns(node, { ...params, styleInfo });

  const attrs = {
    ...extractAlignment(node),
    ...extractParagraphStyles(node),
  };

  // Check if paragraph contains page break
  const hasPageBreak = checkForPageBreak(node);
  if (hasPageBreak) {
    const filteredRuns = runs.filter((run) => run.type !== "hardBreak");
    const paragraphNode: JSONContent = {
      type: "paragraph",
      ...(Object.keys(attrs).length && { attrs }),
      content: filteredRuns.length ? filteredRuns : undefined,
    };
    return [paragraphNode, { type: "horizontalRule" }];
  }

  // Check if pure page break
  if (runs.length === 1 && runs[0].type === "hardBreak") {
    const run = findChild(node, "w:r");
    const br = run && findChild(run, "w:br");
    if (br?.attributes["w:type"] === "page") {
      return { type: "horizontalRule" };
    }
  }

  // Check if pure image
  if (runs.length === 1 && runs[0].type === "image") {
    return runs[0] as JSONContent;
  }

  return {
    type: "paragraph",
    ...(Object.keys(attrs).length && { attrs }),
    content: runs,
  };
}

/**
 * Check if paragraph contains page break
 */
function checkForPageBreak(node: Element): boolean {
  const runElements: Element[] = [];

  const collectRuns = (n: Element) => {
    if (n.name === "w:r") {
      runElements.push(n);
    } else {
      for (const child of n.children) {
        if (child.type === "element") {
          collectRuns(child as Element);
        }
      }
    }
  };

  collectRuns(node);

  return runElements.some((run) => {
    const br = findChild(run, "w:br");
    return br?.attributes["w:type"] === "page";
  });
}

/**
 * Convert to heading (internal function)
 */
async function convertHeading(
  node: Element,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, ImageInfo>;
    options?: DocxImportOptions;
  },
  styleInfo: StyleInfo | undefined,
  level: 1 | 2 | 3 | 4 | 5 | 6,
): Promise<JSONContent> {
  return {
    type: "heading",
    attrs: {
      level,
      ...extractParagraphStyles(node),
    },
    content: await extractRuns(node, { ...params, styleInfo }),
  };
}
