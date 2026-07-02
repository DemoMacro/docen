import { Blockquote as BlockquoteBase } from "@tiptap/extension-blockquote";

/**
 * Blockquote extension — owns the DOCX expression of a blockquote.
 *
 * DOCX has no blockquote element; the convention is a left indent plus a left
 * border on each contained paragraph. This module owns that signature; the
 * DocxManager walks the blockquote's child paragraphs and applies it to each.
 */

/** Left indent (twips, ~0.5 inch) marking a blockquote. */
export const BLOCKQUOTE_INDENT_LEFT = 720;

/** Left border marking a blockquote. */
export const BLOCKQUOTE_BORDER = {
  style: "single",
  size: 18,
  space: 12,
  color: "CCCCCC",
} as const;

/** Apply the blockquote signature (left indent + left border) to paragraph opts. */
export function applyBlockquoteStyle(paraObj: Record<string, unknown>): void {
  const indent = (paraObj.indent as Record<string, unknown> | undefined) ?? {};
  paraObj.indent = { ...indent, left: BLOCKQUOTE_INDENT_LEFT };
  const border = (paraObj.border as Record<string, unknown> | undefined) ?? {};
  paraObj.border = { ...border, left: BLOCKQUOTE_BORDER };
}

// DocxManager applies the signature per child paragraph via applyBlockquoteStyle;
// the extension carries no DOCX attrs of its own.
export { BlockquoteBase as Blockquote };
