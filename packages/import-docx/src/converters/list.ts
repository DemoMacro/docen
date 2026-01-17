import type { Element } from "xast";
import { findChild } from "../utils/xml";

/**
 * Check if a paragraph is a list item
 */
export function isListItem(node: Element): boolean {
  const pPr = findChild(node, "w:pPr");
  return !!pPr && findChild(pPr, "w:numPr") !== undefined;
}

/**
 * Get list numbering info
 */
export function getListInfo(node: Element): {
  numId: string;
  level: number;
} | null {
  const pPr = findChild(node, "w:pPr");
  const numPr = pPr && findChild(pPr, "w:numPr");
  if (!numPr) return null;

  const ilvl = findChild(numPr, "w:ilvl");
  const numId = findChild(numPr, "w:numId");

  if (!ilvl || !numId) return null;

  return {
    numId: numId.attributes["w:val"] as string,
    level: parseInt((ilvl.attributes["w:val"] as string) || "0", 10),
  };
}
