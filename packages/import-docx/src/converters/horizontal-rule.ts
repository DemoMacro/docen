import type { Element } from "xast";

/**
 * Check if a paragraph is a horizontal rule (page break)
 * DOCX represents page breaks as w:br with w:type="page"
 */
export function isHorizontalRule(node: Element): boolean {
  // Check if this paragraph contains only a page break
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:r") {
      const run = child;

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
            const textNode = runChild.children.find((c) => c.type === "text");
            if (textNode && "value" in textNode && textNode.value) {
              const text = (textNode as { value: string }).value.trim();
              if (text.length > 0) {
                hasOtherContent = true;
              }
            }
          } else if (runChild.name !== "w:rPr") {
            // Other elements besides properties
            hasOtherContent = true;
          }
        }
      }

      // Return true if we found a page break and no other content
      if (hasPageBreak && !hasOtherContent) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Helper: Find first child element with given name
 */
function findChild(element: Element, name: string): Element | undefined {
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }
  }
  return undefined;
}
