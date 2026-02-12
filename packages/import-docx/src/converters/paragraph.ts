import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { ParseContext } from "../parser";
import type { StyleInfo } from "../parsers/styles";
import { extractRuns, extractAlignment } from "./text";
import { findChild, parseTwipAttr, convertTwipToCssString } from "@docen/utils";

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
    const left = parseTwipAttr(ind.attributes, "w:left");
    if (left) {
      const leftTwip = parseInt(left, 10);
      result.indentLeft = convertTwipToCssString(leftTwip);
    }

    const right = parseTwipAttr(ind.attributes, "w:right");
    if (right) {
      const rightTwip = parseInt(right, 10);
      result.indentRight = convertTwipToCssString(rightTwip);
    }

    const firstLine = parseTwipAttr(ind.attributes, "w:firstLine");
    if (firstLine) {
      const firstLineTwip = parseInt(firstLine, 10);
      result.indentFirstLine = convertTwipToCssString(firstLineTwip);
    } else {
      // Hanging indent: first line is LESS indented than other lines
      // For example: left=480, hanging=480 means first line indent=0, other lines=480
      // We store the actual first line indent (left - hanging)
      const hanging = parseTwipAttr(ind.attributes, "w:hanging");
      if (hanging) {
        const leftTwip = left ? parseInt(left, 10) : 0;
        const hangingTwip = parseInt(hanging, 10);
        const firstLineTwip = leftTwip - hangingTwip;
        result.indentFirstLine = convertTwipToCssString(firstLineTwip);
      }
    }
  }

  // Extract spacing from w:spacing
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

  return Object.keys(result).length ? result : null;
}

/**
 * Convert DOCX paragraph node to TipTap paragraph
 */
export async function convertParagraph(
  node: Element,
  params: { context: ParseContext; styleInfo?: StyleInfo },
): Promise<JSONContent> {
  const { context, styleInfo: paramStyleInfo } = params;
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const styleName = pStyle?.attributes["w:val"] as string | undefined;

  // Check if it's a heading
  if (styleName && context.styleMap) {
    const styleInfo = context.styleMap.get(styleName);

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

  const styleInfo = styleName && context.styleMap ? context.styleMap.get(styleName) : undefined;
  const runs = await extractRuns(node, { context, styleInfo: paramStyleInfo || styleInfo });

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
    // Wrap image in paragraph to preserve paragraph-level styles (alignment, indents, spacing)
    const imageNode = runs[0] as JSONContent;
    return {
      type: "paragraph",
      ...(Object.keys(attrs).length && { attrs }),
      content: [imageNode],
    };
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
  params: { context: ParseContext },
  styleInfo: StyleInfo | undefined,
  level: 1 | 2 | 3 | 4 | 5 | 6,
): Promise<JSONContent> {
  return {
    type: "heading",
    attrs: {
      level,
      ...extractParagraphStyles(node),
    },
    content: await extractRuns(node, { context: params.context, styleInfo }),
  };
}
