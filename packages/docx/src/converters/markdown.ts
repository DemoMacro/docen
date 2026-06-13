import { MarkdownManager, Markdown } from "@tiptap/markdown";

import type { JSONContent } from "../core";
import { docxExtensions } from "../core";

const markdownManager = new MarkdownManager({
  extensions: [...docxExtensions, Markdown],
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
