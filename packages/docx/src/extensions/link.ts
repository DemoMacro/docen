import type { ParagraphChild, RunOptions } from "@office-open/docx";
import { Link as LinkBase, type LinkOptions } from "@tiptap/extension-link";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";

import { mergeTextNodes } from "../converters/styles";
import type { JSONContent } from "../core";
import { scrollCaretToTop } from "./scroll";
import type { ParseInlineRule, ResolveContext } from "./types";

/**
 * Link — overrides {@link LinkBase}'s click behavior to match MS Word: a plain
 * click places the caret for editing; Ctrl/Cmd+Click follows the hyperlink
 * (in-page scroll for a `#bookmark` anchor, a new tab for an external URL).
 *
 * The upstream extension defaults to `openOnClick: true`, which calls
 * `window.open(href, target)` on every plain click — opening a new tab even for
 * internal `#` anchors. That both breaks TOC navigation and diverges from
 * Word's "click-to-edit / Ctrl+Click-to-follow" semantics.
 */
function docxLinkClickHandler(): Plugin {
  return new Plugin({
    key: new PluginKey("docxLinkClick"),
    props: {
      handleClick(view, _pos, event) {
        const me = event as MouseEvent;
        // Word follows the link only on Ctrl/Cmd + left-click.
        if (me.button !== 0 || !(me.ctrlKey || me.metaKey)) return false;
        const target = me.target as HTMLElement | null;
        const link = target?.closest?.("a");
        if (!link || !view.dom.contains(link)) return false;
        const href = link.getAttribute("href") ?? "";
        me.preventDefault();
        // Internal anchor (#bookmark, e.g. a TOC entry): scroll in-page. The
        // bookmark is exposed as an element id (see InlinePassthrough); posAtDOM
        // maps it to a ProseMirror position so scrollIntoView lands on the
        // heading, since shadow DOM blocks the browser's native #hash jump.
        if (href.startsWith("#")) {
          const id = href.slice(1);
          if (id) {
            const dest = view.dom.querySelector(`[id="${id.replace(/["\\]/g, "\\$&")}"]`);
            if (dest) {
              const p = view.posAtDOM(dest, 0);
              const { state } = view;
              view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, p)));
              // PM's scrollIntoView parks the caret at the bottom edge; scroll it
              // to the top instead (Word-style page follow), matching outline and
              // search-result jumps in the editor.
              scrollCaretToTop(view);
              view.focus();
            }
          }
          return true;
        }
        // External link: open in a new tab (Word follows external hyperlinks).
        if (href) window.open(href, "_blank");
        return true;
      },
    },
  });
}

/** ParagraphChild `{ hyperlink: {...} }` → text[] carrying a link mark. Mirrors
 *  the old DocxManager.resolveHyperlink: recurse the container's runs via ctx,
 *  merge adjacent text, then stamp every text node with the link mark. Returns
 *  null for an empty container or a missing href. */
function resolveHyperlink(
  hyperlink: {
    link?: string;
    anchor?: string;
    tooltip?: string;
    children?: (RunOptions | string)[];
  },
  ctx: ResolveContext,
): JSONContent | null {
  const href = hyperlink.link ?? (hyperlink.anchor ? `#${hyperlink.anchor}` : "");
  if (!href) return null;
  const content = ctx.resolveInlineChildren(
    (hyperlink.children ?? []).map((c) => c as ParagraphChild),
  );
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
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "hyperlink" in child,
  convert: (child, ctx) =>
    resolveHyperlink(
      (
        child as {
          hyperlink: {
            link?: string;
            anchor?: string;
            tooltip?: string;
            children?: (RunOptions | string)[];
          };
        }
      ).hyperlink,
      ctx,
    ),
};

export const Link = LinkBase.extend({
  parseDocxInline,

  addOptions(): LinkOptions {
    // Disable the upstream plain-click window.open (openOnClick); a plain click
    // now places the caret (Word), and docxLinkClickHandler follows on Ctrl+Click.
    // `this.parent?.()` is `LinkOptions | undefined`; spreading widens LinkOptions'
    // required fields to optional in the inferred literal type, so it no longer
    // satisfies LinkOptions even though parent always supplies them at runtime.
    return { ...this.parent?.(), openOnClick: false } as LinkOptions;
  },

  addProseMirrorPlugins() {
    return [...(this.parent?.() ?? []), docxLinkClickHandler()];
  },
});
