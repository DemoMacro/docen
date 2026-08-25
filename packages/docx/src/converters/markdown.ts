import type { MarkdownParseHelpers, MarkdownToken } from "@tiptap/core";
import { MarkdownManager, Markdown } from "@tiptap/markdown";

import { Extension, type JSONContent } from "../core";
import { docxExtensions } from "../core";
import { assignOrderedReferences, HTML_ORDERED_TEMP } from "../extensions/list-numbering";
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

/**
 * Bridge for marked "list" tokens: the schema has no list nodes (a list item
 * is a paragraph with bullet/numbering attrs), so the tree is flattened —
 * each item's text block becomes one list paragraph, nested lists recurse one
 * level deeper. Ordered items carry the html-ordered placeholder reference;
 * parseMarkdown rewrites each run to a fresh generated reference.
 */
const ListTokenBridge = Extension.create({
  name: "listTokenBridge",
  markdownTokenName: "list",
  parseMarkdown: (token: MarkdownToken, h: MarkdownParseHelpers) => {
    const out: JSONContent[] = [];
    const walk = (list: MarkdownToken, level: number): void => {
      for (const item of list.items ?? []) {
        const blocks = (item.tokens ?? []) as MarkdownToken[];
        const text = blocks.find((b) => b.type === "text" || b.type === "paragraph");
        if (text) {
          const attrs = list.ordered
            ? { numbering: { reference: HTML_ORDERED_TEMP, level } }
            : { bullet: { level } };
          out.push(h.createNode("paragraph", attrs, h.parseInline(text.tokens ?? [])));
        }
        for (const block of blocks) {
          if (block.type === "list") walk(block, level + 1);
        }
      }
    };
    walk(token, 0);
    return out;
  },
});

const markdownManager = new MarkdownManager({
  extensions: [...docxExtensions, HeadingTokenBridge, ListTokenBridge, Markdown],
});

/**
 * Parse Markdown string to Tiptap JSON.
 */
export function parseMarkdown(markdown: string): JSONContent {
  return assignOrderedReferences(markdownManager.parse(markdown));
}

/**
 * Generate Markdown string from Tiptap JSON.
 */
export function generateMarkdown(doc: JSONContent): string {
  return markdownManager.serialize(doc);
}
