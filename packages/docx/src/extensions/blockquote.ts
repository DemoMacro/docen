import type { ParagraphOptions } from "@office-open/docx";
import { Blockquote as BlockquoteBase } from "@tiptap/extension-blockquote";

import { cleanAttrs } from "../converters/styles";
import type { JSONContent } from "../core";
import type { ParseAggregatorRule, ResolveContext } from "./types";

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

/** Classify a paragraph as a blockquote member by its signature (left indent
 *  + left border). compile stamps this via applyBlockquoteStyle; this is the
 *  reverse predicate. Pure: reads only the paragraph opts. */
function detectBlockquote(para: ParagraphOptions): boolean {
  const p = para as unknown as Record<string, unknown>;
  const indent = p.indent as { left?: number } | undefined;
  const border = p.border as { left?: Record<string, unknown> } | undefined;
  if (!indent || indent.left !== BLOCKQUOTE_INDENT_LEFT) return false;
  const bl = border?.left;
  if (!bl) return false;
  const sig = BLOCKQUOTE_BORDER;
  return (
    bl.style === sig.style &&
    bl.size === sig.size &&
    bl.space === sig.space &&
    bl.color === sig.color
  );
}

/** Rebuild a blockquote node from a run of signature-carrying paragraphs,
 *  stripping the indent/border signature so child paragraphs render clean. */
function buildBlockquote(group: ParagraphOptions[], ctx: ResolveContext): JSONContent[] {
  const content: JSONContent[] = [];
  for (const para of group) {
    const node = ctx.resolveParagraph(para);
    const attrs = node.attrs as Record<string, unknown> | undefined;
    if (attrs) {
      if (attrs.indent) {
        const indent = { ...(attrs.indent as object) } as Record<string, unknown>;
        delete indent.left;
        attrs.indent = Object.keys(indent).length > 0 ? indent : undefined;
      }
      if (attrs.border) {
        const border = { ...(attrs.border as object) } as Record<string, unknown>;
        delete border.left;
        attrs.border = Object.keys(border).length > 0 ? border : undefined;
      }
      const cleaned = cleanAttrs(attrs);
      if (Object.keys(cleaned).length > 0) node.attrs = cleaned;
      else delete node.attrs;
    }
    content.push(node);
  }
  return [{ type: "blockquote", content }];
}

// DOCX blockquote (left indent + left border signature) → blockquote node.
export const parseDocxAggregator: ParseAggregatorRule = {
  belongs: (para) => detectBlockquote(para),
  build: (group, ctx) => buildBlockquote(group, ctx),
};

// DocxManager applies the signature per child paragraph via applyBlockquoteStyle;
// the extension declares the aggregator so resolve is reflective.
export const Blockquote = BlockquoteBase.extend({ parseDocxAggregator });
