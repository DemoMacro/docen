import {
  generateHTML as generateTiptapHTML,
  generateJSON as generateTiptapJSON,
} from "@tiptap/html";
import type { ParseOptions } from "@tiptap/pm/model";

import type { JSONContent, Extensions } from "../core";
import { docxExtensions } from "../core";

const defaultExtensions: Extensions = docxExtensions;

/**
 * Parse HTML string to Tiptap JSON.
 */
export function parseHTML(
  html: string,
  extensions?: Extensions,
  options?: ParseOptions,
): JSONContent {
  return generateTiptapJSON(html, extensions ?? defaultExtensions, options);
}

/**
 * Generate HTML string from Tiptap JSON.
 */
export function generateHTML(doc: JSONContent, extensions?: Extensions): string {
  return generateTiptapHTML(doc, extensions ?? defaultExtensions);
}
