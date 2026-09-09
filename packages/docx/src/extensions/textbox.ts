import type { SectionChild } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { cleanAttrs } from "../converters/styles";
import { Node } from "../core";
import type { ParseBlockRule } from "./types";
import { attrNative } from "./utils";

/**
 * The legacy VML text box (w:pict > v:shape > v:textbox > w:txbxContent).
 * The box's own data — the structured VML `style` plus the residual paragraph
 * options office-open keeps on the branch for stringify — rides verbatim on
 * attrs.textbox; the contained block stream is editable content, projected
 * into the flow like a content control's. Floating the box at its VML
 * position is a later rendering concern; the text itself is what must be
 * editable and survive the round-trip.
 */

type TextboxBranch = Extract<SectionChild, { textbox: unknown }>;

export const parseDocxBlock: ParseBlockRule<TextboxBranch> = {
  match: (child): child is TextboxBranch => "textbox" in child,
  convert: (child, ctx) => {
    const { children, ...box } = child.textbox;
    const content = ctx.resolveBlockStream((children ?? []) as SectionChild[]);
    if (content.length === 0) content.push({ type: "paragraph" });
    const node: JSONContent = { type: "textbox", content };
    const attrs = cleanAttrs(box as Record<string, unknown>);
    if (Object.keys(attrs).length > 0) node.attrs = { textbox: attrs };
    return node;
  },
};

export const Textbox = Node.create({
  name: "textbox",
  group: "block",
  content: "block+",
  // Box boundaries: isolating keeps Backspace at the start from pulling the
  // first inner paragraph out of the box; defining keeps the shell when the
  // whole content is selected and replaced.
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      textbox: attrNative(),
    };
  },

  parseHTML() {
    return [{ tag: "div.docx-textbox" }];
  },

  parseDocxBlock,
});
