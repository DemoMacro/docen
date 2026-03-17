import {
  generateHTML as generateTiptapHTML,
  generateJSON as generateTiptapJSON,
} from "@tiptap/html/server";
import { tiptapExtensions } from "@docen/extensions";
import type { JSONContent, Extensions } from "@tiptap/core";
import type { ParseOptions } from "@tiptap/pm/model";
import { MarkdownManager } from "@tiptap/markdown";
import { Markdown } from "@tiptap/markdown";

/**
 * Default TipTap extensions used by all converters
 */
const defaultExtensions: Extensions = tiptapExtensions;

/**
 * Shared MarkdownManager instance for Markdown parsing/serialization
 */
const markdownManager = new MarkdownManager({
  extensions: [...tiptapExtensions, Markdown],
});

/**
 * Parse HTML string to TipTap JSON
 */
export function parseHTML(
  html: string,
  extensions?: Extensions,
  options?: ParseOptions,
): JSONContent {
  return generateTiptapJSON(html, extensions ?? defaultExtensions, options);
}

/**
 * Generate HTML string from TipTap JSON
 */
export function generateHTML(doc: JSONContent, extensions?: Extensions): string {
  return generateTiptapHTML(doc, extensions ?? defaultExtensions);
}

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

// Re-export DOCX functions
export { parseDOCX } from "@docen/import-docx";
export { generateDOCX } from "@docen/export-docx";

// Re-export types for convenience
export type { JSONContent } from "@tiptap/core";
export type { Extensions } from "@tiptap/core";
export type { DocxImportOptions } from "@docen/import-docx";
export type { DocxExportOptions } from "@docen/export-docx";
export type * from "@docen/extensions/types";
