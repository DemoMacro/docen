import type { GroupOptions, ParagraphChild } from "@office-open/docx";

import { Node } from "../core";
import { resolveGroupOptions } from "./group-members";
import type { ParseInlineRule } from "./types";

/**
 * wpgGroup — inline node carrying a DOCX drawing group (wpg: wordprocessingGroup)
 * with a modelable interior. The group's own geometry/styling (transformation,
 * chOff/chExt child coordinate space, fill/effects/floating/altText) ride on
 * attrs.wpgGroup; the members (wps shapes, pictures, nested groups) are PM
 * content via group-members. Charts and content parts ride inlinePassthrough
 * member atoms — no PM node of their own, same as the paragraph top level.
 *
 * A group whose payload cannot be modeled (empty/malformed children) resolves
 * to null and falls back to the generic inlinePassthrough — an opaque blob
 * that still round-trips byte-faithfully.
 */

const attrWpgGroup = () => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute("data-wpg-group");
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

// DOCX drawing group (wpg) → editable content node: GroupOptions minus
// children rides on attrs.wpgGroup; children resolve through group-members.
// Null → the generic inlinePassthrough fallback (resolveParagraphChild).
export const parseDocxInline: ParseInlineRule<Extract<ParagraphChild, { wpgGroup: GroupOptions }>> =
  {
    match: (child): child is Extract<ParagraphChild, { wpgGroup: GroupOptions }> =>
      "wpgGroup" in child,
    convert: (child, ctx) => resolveGroupOptions(child.wpgGroup, ctx),
  };

export const WpgGroup = Node.create({
  name: "wpgGroup",
  group: "inline",
  inline: true,
  // Editable member sequence (was an opaque atom). isolating stops Backspace
  // at the border from spilling member edits into the anchor paragraph;
  // defining keeps the node when its interior is fully selected+replaced.
  content: "inline+",
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      wpgGroup: attrWpgGroup(),
    };
  },

  parseHTML() {
    return [{ tag: "span[data-wpg-group]" }];
  },

  parseDocxInline,
});
