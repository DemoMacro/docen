import type { JSONContent } from "@tiptap/core";
import { Paragraph as BaseParagraph } from "@tiptap/extension-paragraph";
import type { Node } from "@tiptap/pm/model";

import { docxParagraphAttrs, renderTextBlock, SECTION_ATTR_KEYS } from "./utils";

/**
 * Paragraph extension with nested office-open attrs.
 *
 * Attrs mirror ParagraphPropertiesOptionsBase (alignment/indent/spacing/border/
 * shading/frame as nested objects + scalar OOXML properties). DOCX round-trip is
 * near-identity: renderDocx/parseDocx pass attrs through; CSS conversion happens
 * only in renderHTML via utils mappers. Heading shares the same attrs (a heading
 * is a paragraph in OOXML) — see docxParagraphAttrs.
 */

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    if (SECTION_ATTR_KEYS.has(key)) continue;
    // styleId (attr) → OOXML `style` (the paragraph's pStyle reference).
    if (key === "styleId") {
      opts.style = value;
      continue;
    }
    opts[key] = value;
  }
  return opts;
}

/**
 * Structural/semantic keys handled elsewhere (heading ext, list handling, run/text
 * children). NOTE: `run` is intentionally NOT skipped — ParagraphOptions.run (the
 * paragraph's default run properties: font/size/color) is kept as an attr for
 * lossless round-trip (e.g. header/footer paragraphs whose styling lives there).
 */
const SKIP_KEYS = new Set([
  "heading",
  "style",
  "bullet",
  "numbering",
  "children",
  "text",
  "thematicBreak",
]);

export function parseDocx(opts: Record<string, unknown>): Record<string, unknown> {
  const resolved = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};
  // OOXML `style` (the paragraph's pStyle reference, e.g. "Heading1") → styleId,
  // carried as an attr so the named style's CSS applies via class="docx-style-{id}".
  if (resolved.style) attrs.styleId = resolved.style;
  for (const [key, value] of Object.entries(resolved)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  return attrs;
}

// ── Extension ──

export const Paragraph = BaseParagraph.extend({
  // A heading is a paragraph in OOXML (a <w:p> with pStyle="Heading1"), so
  // Paragraph and Heading share the SAME office-open paragraph attrs via
  // docxParagraphAttrs — only Heading adds `level`. See utils.
  addAttributes() {
    return { ...this.parent?.(), ...docxParagraphAttrs() };
  },

  renderHTML({ node, HTMLAttributes }: { node: Node; HTMLAttributes: Record<string, unknown> }) {
    return renderTextBlock(node, HTMLAttributes, "p");
  },

  renderDocx,
  parseDocx,
});
