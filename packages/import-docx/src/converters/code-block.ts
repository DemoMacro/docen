import type { Element } from "xast";
import { findChild } from "@docen/utils";

/**
 * Check if a paragraph is a code block
 */
export function isCodeBlock(node: Element): boolean {
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const style = pStyle?.attributes["w:val"] as string;

  return style === "CodeBlock" || style?.startsWith("Code") || false;
}

/**
 * Get code block language
 */
export function getCodeBlockLanguage(node: Element): string | undefined {
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const style = pStyle?.attributes["w:val"] as string;

  if (!style?.startsWith("CodeBlock")) return undefined;

  const lang = style.replace("CodeBlock", "").toLowerCase();
  return lang || undefined;
}
