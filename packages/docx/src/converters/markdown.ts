import type { MarkdownParseHelpers, MarkdownToken } from "@tiptap/core";
import { MarkdownManager, Markdown } from "@tiptap/markdown";

import { Extension, type JSONContent } from "../core";
import { docxExtensions } from "../core";
import { HEADING_COMPILE_MAP } from "../extensions/paragraph";

/**
 * Bridge for marked "heading" tokens: the schema has no heading node (a
 * heading is a paragraph with a HeadingLevel attr), so the markdown manager's
 * fallback would emit a `heading` node type schema.nodeFromJSON rejects. This
 * extension claims the token and produces the paragraph form instead. It adds
 * no schema — the manager reads only the parseMarkdown config field.
 */
const HeadingTokenBridge = Extension.create({
  name: "headingTokenBridge",
  markdownTokenName: "heading",
  parseMarkdown: (token: MarkdownToken, h: MarkdownParseHelpers) =>
    h.createNode(
      "paragraph",
      { heading: HEADING_COMPILE_MAP[token.depth ?? 1] ?? "Heading1" },
      h.parseInline(token.tokens ?? []),
    ),
});

const markdownManager = new MarkdownManager({
  extensions: [...docxExtensions, HeadingTokenBridge, Markdown],
});

/**
 * Parse Markdown string to Tiptap JSON.
 */
export function parseMarkdown(markdown: string): JSONContent {
  return markdownManager.parse(markdown);
}

/**
 * Generate Markdown string from Tiptap JSON.
 */
export function generateMarkdown(doc: JSONContent): string {
  return markdownManager.serialize(doc);
}
