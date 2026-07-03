import type { ParagraphChild, ParagraphOptions, RunOptions } from "@office-open/docx";
import { CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";

import type { JSONContent } from "../core";
import type { ParseParagraphRule, ResolveContext } from "./types";

/**
 * CodeBlock extension — CodeBlockLowlight with a DOCX renderDocx + parse rule.
 *
 * DOCX has no dedicated code-block element. A code block maps to a single
 * paragraph styled "Code" with a monospace run font. Inline marks (syntax-
 * highlight tokens) and line breaks (`\n` → `<w:br/>`) are handled by
 * DocxManager's shared inline-content compilation; resolveCodeBlock reassembles
 * the text. This module owns the paragraph-level style + the parse rule that
 * recognizes a "Code"-styled paragraph.
 *
 * `language` has no OOXML carrier and is intentionally dropped (known lossy).
 */

/** codeBlock node → paragraph properties (style "Code" + monospace font). */
export function renderDocx(_node: JSONContent): Record<string, unknown> {
  return { style: "Code", run: { font: "Consolas" } };
}

/** ParagraphOptions styled "Code" → codeBlock node. Reassembles code: a break
 *  child → "\n" (merged into the previous text node), a text child keeps its
 *  run marks (syntax-highlight tokens). Mirrors the old resolveCodeBlock. */
function resolveCodeBlock(opts: ParagraphOptions, ctx: ResolveContext): JSONContent {
  const children = opts.children as (ParagraphChild | string)[] | undefined;
  const content: JSONContent[] = [];
  if (children) {
    for (const child of children) {
      if (typeof child === "string") {
        if (child) content.push({ type: "text", text: child });
      } else if (typeof child === "object" && child !== null) {
        if ("break" in child) {
          const prev = content[content.length - 1];
          if (prev && prev.type === "text") prev.text = (prev.text ?? "") + "\n";
          else content.push({ type: "text", text: "\n" });
        } else if ("text" in child) {
          const marks = ctx.resolveMarks(child as RunOptions);
          const textNode: JSONContent = { type: "text", text: (child as { text: string }).text };
          if (marks) textNode.marks = marks;
          content.push(textNode);
        }
      }
    }
  } else if (opts.text) {
    content.push({ type: "text", text: opts.text });
  }
  const node: JSONContent = { type: "codeBlock" };
  if (content.length > 0) node.content = content;
  return node;
}

// DOCX paragraph styled "Code" → codeBlock node (resolveCodeBlock reassembles
// the text from runs + breaks).
export const parseDocxParagraph: ParseParagraphRule = {
  match: (para) => para.style === "Code",
  convert: (para, ctx) => resolveCodeBlock(para, ctx),
};

export const CodeBlock = CodeBlockLowlight.extend({ renderDocx, parseDocxParagraph });
