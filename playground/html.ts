import {
  generateHTML as generateTiptapHTML,
  generateJSON as generateTiptapJSON,
} from "@tiptap/html";
import { tiptapExtensions } from "./extensions";
import { JSONContent, Extensions } from "@tiptap/core";
import { ParseOptions } from "@tiptap/pm/model";

export function generateHTML(doc: JSONContent, extensions?: Extensions) {
  return generateTiptapHTML(doc, extensions ?? tiptapExtensions);
}

export function generateJSON(
  html: string,
  extensions?: Extensions,
  options?: ParseOptions,
): JSONContent {
  return generateTiptapJSON(html, extensions ?? tiptapExtensions, options);
}
