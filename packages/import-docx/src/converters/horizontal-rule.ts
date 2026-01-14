import type { Element, Text } from "xast";
import { findChild } from "../utils/xml";

/**
 * Check if a paragraph is a horizontal rule (page break)
 * DOCX represents page breaks as w:br with w:type="page"
 */
export function isHorizontalRule(node: Element): boolean {
  // Check if this paragraph contains only a page break
  const run = findChild(node, "w:r");
  if (!run) return false;

  // Check all children of the run
  let hasPageBreak = false;
  let hasOtherContent = false;

  for (const runChild of run.children) {
    if (runChild.type === "element") {
      if (runChild.name === "w:br") {
        const brType = runChild.attributes["w:type"];
        if (brType === "page") {
          hasPageBreak = true;
        }
      } else if (runChild.name === "w:t") {
        // Check if text element has content
        const textNode = runChild.children.find((c): c is Text => c.type === "text");
        if (textNode && textNode.value && textNode.value.trim().length > 0) {
          hasOtherContent = true;
        }
      } else if (runChild.name !== "w:rPr") {
        // Other elements besides properties
        hasOtherContent = true;
      }
    }
  }

  // Return true if we found a page break and no other content
  return hasPageBreak && !hasOtherContent;
}
