import type { Element } from "xast";
import { findChild, DOCX_STYLE_NAMES } from "@docen/utils";

/**
 * Check if a paragraph is a code block
 */
export function isCodeBlock(node: Element): boolean {
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const style = pStyle?.attributes["w:val"] as string;

  return (
    style === DOCX_STYLE_NAMES.CODE_BLOCK ||
    style?.startsWith(DOCX_STYLE_NAMES.CODE_PREFIX) ||
    false
  );
}

/**
 * Get code block language
 */
export function getCodeBlockLanguage(node: Element): string | undefined {
  const pPr = findChild(node, "w:pPr");
  const pStyle = pPr && findChild(pPr, "w:pStyle");
  const style = pStyle?.attributes["w:val"] as string;

  if (!style?.startsWith(DOCX_STYLE_NAMES.CODE_BLOCK)) return undefined;

  const lang = style.replace(DOCX_STYLE_NAMES.CODE_BLOCK, "").toLowerCase();
  return lang || undefined;
}
