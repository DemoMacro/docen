import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "../option";
import { extractRuns, extractAlignment } from "./text";
import { findChild } from "../utils/xml";

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
  return {
    type: "paragraph",
    ...(attrs && { attrs }),
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
  return {
    type: "heading",
    attrs: { level },
    content: await extractRuns(node, params),
  };
}
