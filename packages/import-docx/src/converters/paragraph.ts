import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import { extractRuns, extractAlignment } from "./text";

/**
 * Convert DOCX paragraph node to TipTap paragraph
 */
export function convertParagraph(
  node: Element,
  hyperlinks: Map<string, string>,
  images: Map<string, string>,
): JSONContent {
  // Check if it's a heading by finding w:pPr > w:pStyle
  let styleName: string | undefined;
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:pPr") {
      const pPr = child;
      for (const pPrChild of pPr.children) {
        if (pPrChild.type === "element" && pPrChild.name === "w:pStyle") {
          styleName = pPrChild.attributes["w:val"] as string;
          break;
        }
      }
      break;
    }
  }

  if (styleName) {
    // Check if it's a heading (e.g., "Heading1", "Heading2")
    const headingMatch = styleName.match(/^Heading(\d)$/);
    if (headingMatch) {
      const level = parseInt(headingMatch[1]) as 1 | 2 | 3 | 4 | 5 | 6;
      return convertHeading(node, hyperlinks, level, images);
    }
  }

  // Extract runs (text, images, hardBreaks)
  const runs = extractRuns(hyperlinks, node, images);

  // Check if this is a horizontal rule (page break)
  if (runs.length === 1 && runs[0].type === "hardBreak") {
    // Check if it's a page break type
    for (const child of node.children) {
      if (child.type === "element" && child.name === "w:r") {
        for (const rChild of child.children) {
          if (rChild.type === "element" && rChild.name === "w:br") {
            if (rChild.attributes["w:type"] === "page") {
              return { type: "horizontalRule" };
            }
          }
        }
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
function convertHeading(
  node: Element,
  hyperlinks: Map<string, string>,
  level: 1 | 2 | 3 | 4 | 5 | 6,
  images: Map<string, string>,
): JSONContent {
  return {
    type: "heading",
    attrs: { level },
    content: extractRuns(hyperlinks, node, images),
  };
}
