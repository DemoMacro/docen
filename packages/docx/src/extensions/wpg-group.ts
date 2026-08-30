import type { GroupOptions, ParagraphChild, ShapeCoreOptions } from "@office-open/docx";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";

/**
 * wpgGroup — inline atom carrying a DOCX drawing group (wpg: wordprocessingGroup)
 * as an opaque blob. Mirrors the office-open `GroupOptions` / ParagraphChild
 * `wpgGroup` field verbatim in attrs.wpgGroup, so the node name and attr stay
 * aligned with the OOXML concept (CT_WordprocessingGroup).
 *
 * A group bundles pictures (pic), shapes (wps), and nested groups (wpg) behind a
 * shared coordinate space (grpSpPr chOff/chExt → extent). Rendering is owned by
 * the editor layer's NodeView, which lays each child out at its transformed
 * position/size so the group appears as Word draws it.
 */

/** wps interior data (fill/outline/bodyProperties/text body). */
export type WpsData = ShapeCoreOptions;

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

// DOCX drawing group (wpg) → opaque atom: full GroupOptions rides on
// attrs.wpgGroup (the editor doesn't model the group interior).
export const parseDocxInline: ParseInlineRule<Extract<ParagraphChild, { wpgGroup: GroupOptions }>> =
  {
    match: (child): child is Extract<ParagraphChild, { wpgGroup: GroupOptions }> =>
      "wpgGroup" in child,
    convert: (child) => ({
      type: "wpgGroup",
      attrs: { wpgGroup: child.wpgGroup },
    }),
  };

export const WpgGroup = Node.create({
  name: "wpgGroup",
  group: "inline",
  inline: true,
  atom: true,

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
