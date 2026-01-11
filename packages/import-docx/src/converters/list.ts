import type { Element } from "xast";

/**
 * Check if a paragraph is a list item
 */
export function isListItem(node: Element): boolean {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return false;

  const numPr = findChild(pPr, "w:numPr");
  return !!numPr;
}

/**
 * Get list numbering info
 */
export function getListInfo(node: Element): {
  numId: string;
  level: number;
} | null {
  const pPr = findChild(node, "w:pPr");
  if (!pPr) return null;

  const numPr = findChild(pPr, "w:numPr");
  if (!numPr) return null;

  const ilvl = findChild(numPr, "w:ilvl");
  const numId = findChild(numPr, "w:numId");

  if (!ilvl || !numId) return null;

  return {
    numId: numId.attributes["w:val"] as string,
    level: parseInt((ilvl.attributes["w:val"] as string) || "0"),
  };
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
