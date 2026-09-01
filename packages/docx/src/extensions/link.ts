import type { ParagraphChild } from "@office-open/docx";
import { Link as LinkBase, type LinkOptions } from "@tiptap/extension-link";

import { mergeTextNodes } from "../converters/styles";
import type { JSONContent } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/** The inline hyperlink ParagraphChild branch, derived from the union (the
 *  payload type is not exported separately). */
type HyperlinkBranch = Extract<ParagraphChild, { hyperlink: unknown }>;

/**
 * Link — DOCX hyperlink conversion on top of {@link LinkBase}. Click-to-follow
 * lives in the editor's canvas layer (there is no DOM view here to click); the
 * engine only turns `w:hyperlink` into link marks and back.
 *
 * The upstream defaults `openOnClick`/`autolink` stay off so a plain click
 * never `window.open`s, and a typed URL never silently gains a link mark
 * without the "Hyperlink" character style the insert paths stamp.
 */

/** ParagraphChild `{ hyperlink: {...} }` → text[] carrying a link mark. Mirrors
 *  the old DocxManager.resolveHyperlink: recurse the container's runs via ctx,
 *  merge adjacent text, then stamp every text node with the link mark. Returns
 *  null for an empty container or a missing href. */
function resolveHyperlink(
  hyperlink: HyperlinkBranch["hyperlink"],
  ctx: ResolveContext,
): JSONContent | null {
  const href = hyperlink.url ?? (hyperlink.anchor ? `#${hyperlink.anchor}` : "");
  if (!href) return null;
  // Hyperlink children (runs, strings, nested branches) are all valid inline
  // input — the ParagraphChild union admits RunOptions as its fallback member.
  const content = ctx.resolveInlineChildren(hyperlink.children ?? []);
  if (content.length === 0) return null;
  const merged = mergeTextNodes(content);
  for (const node of merged) {
    if (node.type === "text") {
      node.marks = [
        ...(node.marks ?? []),
        {
          type: "link",
          attrs: {
            href,
            // Internal anchor (#bookmark, e.g. a TOC entry jump) stays in-window
            // so the in-page scroll resolves; only external links open a tab.
            target: href.startsWith("#") ? null : "_blank",
            rel: "noopener noreferrer nofollow",
            class: null,
            title: hyperlink.tooltip ?? null,
          },
        },
      ];
    }
  }
  return merged;
}

// DOCX hyperlink run → office-open ParagraphChild `{ hyperlink: {...} }`.
export const parseDocxInline: ParseInlineRule<HyperlinkBranch> = {
  match: (child): child is HyperlinkBranch => "hyperlink" in child,
  convert: (child, ctx) => resolveHyperlink(child.hyperlink, ctx),
};

export const Link = LinkBase.extend({
  parseDocxInline,

  addOptions(): LinkOptions {
    // `this.parent?.()` is `LinkOptions | undefined`; spreading widens LinkOptions'
    // required fields to optional in the inferred literal type, so it no longer
    // satisfies LinkOptions even though parent always supplies them at runtime.
    return { ...this.parent?.(), openOnClick: false, autolink: false } as LinkOptions;
  },
});
