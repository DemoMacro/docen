import type { JSONContent } from "@tiptap/core";

/**
 * Parse plain text string to TipTap JSON
 */
export function parseText(text: string): JSONContent {
  return {
    type: "doc",
    content: text
      .split("\n")
      .map((line) => ({
        type: "paragraph",
        content: line ? [{ type: "text", text: line }] : [],
      })),
  };
}

/**
 * Generate plain text string from TipTap JSON
 */
export function generateText(doc: JSONContent): string {
  function walk(node: JSONContent): string {
    if (node.text) return node.text;
    if (!node.content) return "";
    return node.content.map(walk).join("");
  }

  return walk(doc);
}
