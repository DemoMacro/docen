import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { ParseContext } from "../parser";
import type { StyleInfo } from "../parsers/styles";
import { extractRuns, extractAlignment } from "./text";
import { findChild } from "@docen/utils";
import { extractParagraphStyles, resolveStyleInfo } from "../parsers/styles";

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

  // First check for direct outlineLvl setting (higher priority than style-based)
  if (pPr) {
    const outlineLvlElement = findChild(pPr, "w:outlineLvl");
    if (outlineLvlElement?.attributes["w:val"] !== undefined) {
      const outlineLvl = parseInt(outlineLvlElement.attributes["w:val"] as string, 10);
      if (outlineLvl >= 0 && outlineLvl <= 5) {
        const level = (outlineLvl + 1) as 1 | 2 | 3 | 4 | 5 | 6;
        const styleInfo = context.styleMap
          ? resolveStyleInfo(context.styleMap, styleName)
          : undefined;
        return convertHeading(node, params, styleInfo, level);
      }
    }
  }

  // Then check style-based heading detection
  if (styleName && context.styleMap) {
    const styleInfo = resolveStyleInfo(context.styleMap, styleName);

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

  const styleInfo = context.styleMap ? resolveStyleInfo(context.styleMap, styleName) : undefined;
  const runs = await extractRuns(node, { context, styleInfo: paramStyleInfo || styleInfo });

  const attrs = {
    ...extractAlignment(node),
    ...extractParagraphStyles(node, styleInfo),
  };

  // Check if pure page break (must check before general page break handler)
  if (runs.length === 1 && runs[0].type === "hardBreak") {
    const run = findChild(node, "w:r");
    const br = run && findChild(run, "w:br");
    if (br?.attributes["w:type"] === "page") {
      return { type: "horizontalRule" };
    }
  }

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
      ...extractParagraphStyles(node, styleInfo),
    },
    content: await extractRuns(node, { context: params.context, styleInfo }),
  };
}
