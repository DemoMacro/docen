import {
  generateHTML as generateTiptapHTML,
  generateJSON as generateTiptapJSON,
} from "@tiptap/html";
import { tiptapExtensions } from "@docen/extensions";
import type { JSONContent, Extensions } from "@tiptap/core";
import type { ParseOptions } from "@tiptap/pm/model";

/**
 * Default TipTap extensions used by all converters
 */
const defaultExtensions: Extensions = tiptapExtensions;

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
