import { Node } from "../core";
import type { ParseInlineRule } from "./types";

/**
 * Tab — an inline atom representing a DOCX `<w:r><w:tab/></w:r>` tab character.
 *
 * office-open parses `<w:tab/>` as `{ tab: true }` inside a run's children. The
 * resolve path turns that into this node so the tab is not lost: previously
 * `<w:tab/>` was dropped, which let `mergeTextNodes` collapse a TOC entry's title
 * and page number into adjacent text (no leader, no right alignment). The node is
 * zero-width, non-editable, and carries no text height — measure skips it the
 * same way it skips other inline atoms, so pagination is unaffected. It only
 * marks where a tab leader (e.g. a TOC's dotted leader) renders. compile turns it
 * back into `{ tab: true }`.
 */

// `<w:r><w:tab/></w:r>` → office-open ParagraphChild `{ tab: true }`. Turned into
// this atom so the tab is not lost (mergeTextNodes would otherwise collapse a
// TOC entry's title and page number together).
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "tab" in child,
  convert: () => ({ type: "tab" }),
};

export const Tab = Node.create({
  name: "tab",
  group: "inline",
  inline: true,
  atom: true,

  parseHTML() {
    return [{ tag: "span.docx-tab" }];
  },

  renderHTML() {
    return ["span", { class: "docx-tab", contenteditable: "false" }];
  },

  parseDocxInline,
});
