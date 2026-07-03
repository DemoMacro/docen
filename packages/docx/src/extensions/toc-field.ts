import type { TableOfContentsOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { cleanAttrs } from "../converters/styles";
import { Node } from "../core";
import type { ParseBlockRule, ResolveContext } from "./types";
import { attrNative } from "./utils";

/**
 * TOC field (`tocField`) — a block container representing a DOCX table of
 * contents, with the rendered entries as editable `content` and the TOC field
 * switches on `attrs.options`.
 *
 * Named `tocField` (not `tableOfContents`) to avoid colliding with the official
 * `@tiptap/extension-table-of-contents`, which is an `Extension` (a live outline
 * generator injected into the editor) that already owns the `tableOfContents`
 * name. Same-name extensions dedupe in Tiptap, and since the official one is not
 * a node it would erase this node type from the schema. The two coexist for
 * different purposes: this node persists a DOCX's rendered TOC; the official
 * extension drives the heading-outline pane.
 *
 * Structuring the TOC as a node (instead of opaque passthrough) is what fixes
 * the export crash. Each entry paragraph's `w:hyperlink` wraps a HYPERLINK field
 * whose content-less runs (fldChar begin/separate/end) office-open parses as
 * `null`. As opaque passthrough those nulls survived verbatim to
 * `generateDocument`, where office-open's `stringifyRunInline(null).break`
 * crashed. The `parseDocxBlock` rule resolves the entries through the shared
 * block-stream path, which drops the nulls (the `child !== null` guard), so
 * compile rebuilds clean entries and the generate path never sees a null — no
 * office-open change required.
 */

// ── Block parse rule (resolve: toc → tocField node) ──

/**
 * Declarative block parse rule: recognize a table of contents SectionChild and
 * rebuild it as an editable `tocField` container. DocxManager dispatches every
 * SectionChild through this rule before the paragraph/passthrough fallbacks. */
export const parseDocxBlock: ParseBlockRule = {
  match: (child) => "toc" in child,
  convert: (child, ctx) =>
    resolveToc((child as { toc: TableOfContentsOptions & { alias?: string } }).toc, ctx),
};

/** Resolve a table of contents into an editable `tocField` container:
 *  `attrs.options` carries the field switches, `content` is the entry
 *  paragraphs. Each entry's inner HYPERLINK field has content-less runs that
 *  office-open parses as `null`; resolving the entries through the shared
 *  block-stream path drops those nulls (the `stringifyRunInline(null).break`
 *  crash). When `entries` is absent/empty (a fresh, unrendered TOC), keep the
 *  node valid for `content: "block+"` with a placeholder empty paragraph. */
function resolveToc(
  toc: TableOfContentsOptions & { alias?: string },
  ctx: ResolveContext,
): JSONContent {
  const { entries, ...options } = toc;
  const content: JSONContent[] = [];
  for (const entry of entries ?? []) {
    const node = ctx.resolveBlock(entry);
    if (!node) continue;
    if (Array.isArray(node)) content.push(...node);
    else content.push(node);
  }
  if (content.length === 0) content.push({ type: "paragraph" });
  const node: JSONContent = { type: "tocField", content };
  const cleanOptions = cleanAttrs(options as Record<string, unknown>);
  if (Object.keys(cleanOptions).length > 0) node.attrs = { options: cleanOptions };
  return node;
}

export const TocField = Node.create({
  name: "tocField",
  group: "block",
  content: "block+",

  addAttributes() {
    return {
      // TOC field switches (hyperlink, headingStyleRange, …); carried verbatim,
      // never rendered to HTML.
      options: attrNative(),
    };
  },

  parseHTML() {
    return [{ tag: "div.docx-toc" }];
  },

  renderHTML() {
    return ["div", { class: "docx-toc" }, 0];
  },

  parseDocxBlock,
});
