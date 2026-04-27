import { tiptapExtensions } from "@docen/extensions";
import type { JSONContent } from "@tiptap/core";
import { MarkdownManager } from "@tiptap/markdown";
import { Markdown } from "@tiptap/markdown";

/**
 * Shared MarkdownManager instance for Markdown parsing/serialization
 */
const markdownManager = new MarkdownManager({
  extensions: [...tiptapExtensions, Markdown],
});

/**
 * Parse Markdown string to TipTap JSON
 */
export function parseMarkdown(markdown: string): JSONContent {
  return markdownManager.parse(markdown);
}

/**
 * Generate Markdown string from TipTap JSON
 */
export function generateMarkdown(doc: JSONContent): string {
  return markdownManager.serialize(doc);
}
