import type { DOMOutputSpec } from "@tiptap/pm/model";

import { Node } from "../core";
import { wpsShapeStyles, type WpsShapeStandalone } from "./wpg-group";

/**
 * wpsShape — inline node carrying a standalone DOCX text-box shape
 * (wp:anchor > wps:wsp > wps:txbx; NOT inside a wpg group). The shape geometry
 * + styling (transformation/floating/fill/outline/bodyProperties) ride on
 * attrs.wpsShape; the editable text body is PM content (block+), one paragraph
 * per office-open ParagraphOptions. Unlike a group's interior wps children
 * (laid out in the group's coordinate space), this shape floats on its own
 * anchor. The engine node has no NodeView (UI-free); the editor layer extends
 * it with a two-element NodeView (outer placement/rotation, inner contentDOM).
 */

const attrWpsShape = () => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute("data-wps-shape");
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

export const WpsShape = Node.create({
  name: "wpsShape",
  group: "inline",
  inline: true,
  // Editable text body (was atom). content:"block+" holds the textbox's
  // paragraph(s); isolating stops Backspace at the start from merging the first
  // paragraph back into the anchor paragraph; defining keeps the node when the
  // body is fully selected+replaced.
  content: "block+",
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      wpsShape: attrWpsShape(),
    };
  },

  parseHTML() {
    return [
      {
        tag: "div[data-wps-shape]",
        // The editable body lives in the inner <div> (the contentDOM); parse it
        // from there (querySelector "div" = the first child div) instead of the
        // outer positioning wrapper, so paragraphs resolve as content rather
        // than getting hoisted out of the inline node.
        contentElement: "div",
      },
    ];
  },

  renderHTML({
    node,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const ws = (node.attrs.wpsShape as WpsShapeStandalone | null) ?? {};
    const { outer, inner, paragraphAnchor } = wpsShapeStyles(ws);
    // Serialize the shape geometry so generateHTML→parseHTML round-trips it
    // (parseHTML JSON.parses data-wps-shape; "" would throw and drop it). The
    // text body is NOT serialized here — it round-trips as PM content.
    const attrs: Record<string, string> = {
      "data-wps-shape": JSON.stringify(ws),
      style: outer,
    };
    if (paragraphAnchor) attrs["data-float-anchor"] = "paragraph";
    return ["div", attrs, ["div", { style: inner }, 0]] as unknown as DOMOutputSpec;
  },
});
