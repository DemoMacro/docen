import type { Element } from "xast";

/**
 * Check if a paragraph is a code block
 */
export function isCodeBlock(node: Element): boolean {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return false;

  const pStyle = findChild(pPr, "w:pStyle");
  if (!pStyle) return false;

  const style = pStyle.attributes["w:val"] as string;
  // Check for code block style patterns
  return style === "CodeBlock" || style?.startsWith("Code");
}

/**
 * Get code block language
 */
export function getCodeBlockLanguage(node: Element): string | undefined {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return undefined;

  const pStyle = findChild(pPr, "w:pStyle");
  if (!pStyle) return undefined;

  const style = pStyle.attributes["w:val"] as string;
  // Extract language from style name like "CodeBlockJavaScript" -> "javascript"
  if (style?.startsWith("CodeBlock")) {
    const lang = style.replace("CodeBlock", "").toLowerCase();
    return lang || undefined;
  }

  return undefined;
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
