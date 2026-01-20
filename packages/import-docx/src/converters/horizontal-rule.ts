import type { Element, Text } from "xast";
import { findChild } from "@docen/utils";

/**
 * Check if a paragraph is a horizontal rule (page break)
 */
export function isHorizontalRule(node: Element): boolean {
  const run = findChild(node, "w:r");
  if (!run) return false;

  let hasPageBreak = false;
  let hasOtherContent = false;

  for (const runChild of run.children) {
    if (runChild.type !== "element") continue;

    if (runChild.name === "w:br" && runChild.attributes["w:type"] === "page") {
      hasPageBreak = true;
    } else if (runChild.name === "w:t") {
      const textNode = runChild.children.find((c): c is Text => c.type === "text");
      if (textNode?.value?.trim().length) {
        hasOtherContent = true;
      }
    } else if (runChild.name !== "w:rPr") {
      hasOtherContent = true;
    }
  }

  return hasPageBreak && !hasOtherContent;
}
