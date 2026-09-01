import type { JSONContent, MarkdownRendererHelpers } from "@tiptap/core";
import { Node } from "@tiptap/core";

import { attrNative } from "./utils";

/**
 * Document extension carrying DOCX document-level data through the Tiptap JSON
 * for lossless round-trip (declared as attrs so editor setContent → getJSON
 * preserves them, not just the standalone converters):
 *
 * - `attrs.styles` — office-open `StylesOptions` (styles.xml: importedStyles /
 *   docDefaultsXml / latentStylesXml as raw XML).
 * - `attrs.core` — docProps/core.xml properties (title/creator/description/…,
 *   see DocxCoreProperties in converters/docx.ts).
 * - `attrs.sectionProperties` — the last section's page layout (page size/margin/
 *   orientation, columns, type, grid; intermediate sections carry theirs on
 *   sectionBreak nodes).
 *
 * None of these render anywhere — they ride the JSON for the converters and
 * the canvas editor's projection to consume.
 *
 * Factory form (`createDocument`): the editor layer may parameterize a
 * different top-level `content` expression but keeps the SAME DOCX attrs.
 * Building it via this factory keeps the Document definition in ONE place
 * (here), instead of `.extend`-overriding this Document and re-stating the
 * attrs. `Document` is the default flat `doc > block+` shape used by the docx
 * package itself.
 */

export function createDocument(content = "block+") {
  return Node.create({
    name: "doc",
    content,
    addAttributes() {
      return {
        styles: attrNative(),
        core: attrNative(),
        sectionProperties: attrNative(),
        sectionHeaders: attrNative(),
        sectionFooters: attrNative(),
        background: attrNative(),
        documentExtras: attrNative(),
        // Source numbering.config (abstractNum definitions) carried verbatim so
        // list markers (glyph/font/indent) round-trip; compile merges it with
        // any regenerated ordered-list definitions.
        numbering: attrNative(),
      };
    },

    // Markdown serialization with flat-list grouping: consecutive list
    // paragraphs (bullet/numbering attrs) render as Markdown list items —
    // "- item" / "N. item" at the flat depth (4-space indent per level, safe
    // under both "- " and "N. " markers). Ordered counters advance per
    // reference and reset deeper levels, mirroring the canvas numbering
    // resolver (layout/project.ts). Non-list children render via the default
    // per-node path, so heading/empty-paragraph semantics stay with Paragraph.
    renderMarkdown: (node: JSONContent, h: MarkdownRendererHelpers): string => {
      type ListAttrs = {
        bullet?: { level?: number } | null;
        numbering?: { reference?: string; level?: number } | null;
      };
      const parts: string[] = [];
      let items: string[] = [];
      const counters = new Map<string, number[]>();
      const flushItems = (): void => {
        if (items.length > 0) {
          parts.push(items.join("\n"));
          items = [];
        }
      };
      (node.content ?? []).forEach((child, index) => {
        const attrs = (child.attrs ?? {}) as ListAttrs;
        const level = attrs.bullet?.level ?? attrs.numbering?.level ?? 0;
        const reference = attrs.numbering
          ? (attrs.numbering.reference ?? "")
          : attrs.bullet
            ? "bullet"
            : null;
        if (reference == null) {
          flushItems();
          parts.push(h.renderChild ? h.renderChild(child, index) : h.renderChildren([child]));
          return;
        }
        const indent = " ".repeat(4 * level);
        let marker: string;
        if (attrs.bullet) {
          marker = "-";
        } else {
          const counts = counters.get(reference) ?? [];
          counters.set(reference, counts);
          counts[level] = (counts[level] ?? 0) + 1;
          counts.length = level + 1;
          marker = `${counts[level]}.`;
        }
        const body = h.renderChild ? h.renderChild(child, index) : h.renderChildren([child]);
        items.push(`${indent}${marker} ${body.split("\n").join(`\n${indent}`)}`.trimEnd());
      });
      flushItems();
      return parts.join("\n\n");
    },
  });
}

/** Default flat Document (`doc > block+`) — the DOCX round-trip shape. */
export const Document = createDocument();
