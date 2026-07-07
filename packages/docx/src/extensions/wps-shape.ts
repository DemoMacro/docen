import type { ParagraphOptions } from "@office-open/docx";
import type { DOMOutputSpec } from "@tiptap/pm/model";

import { cleanAttrs } from "../converters/styles";
import type { JSONContent } from "../core";
import { Node } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";
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

/** ParagraphChild `{ wpsShape: {...} }` → wpsShape node. Mirrors the old
 *  DocxManager wpsShape branch: the shape's text body (children) becomes PM
 *  content (one node per paragraph); geometry/styling ride on attrs.wpsShape.
 *  Each paragraph's defRPr (para.run) is merged into its runs then dropped — it
 *  is the box's default run-properties, not the ¶-mark rPr (see inline note). */
function resolveWpsShape(
  ws: { children?: (ParagraphOptions | string)[] } & Record<string, unknown>,
  ctx: ResolveContext,
): JSONContent {
  const content: JSONContent[] = [];
  if (ws?.children) {
    for (const para of ws.children) {
      if (typeof para !== "object" || para === null) {
        const node = ctx.resolveParagraph(para);
        if (node) content.push(node);
        continue;
      }
      // DrawingML defRPr (para.run) is the default run-properties for the box's
      // runs, NOT the OOXML ¶-mark rPr. Merge it into each run (matching the
      // prior atom renderWpsText: {...para.run, ...r}), then drop it from the
      // paragraph (run: undefined): paragraph.ts renders attrs.run.size as
      // ¶-mark line-height, which would override the box's grid line-height —
      // but defRPr is a run default, not a ¶ mark. Round-trip safe — runs carry
      // the full rPr, so compile emits per-run rPr and Word renders identically.
      const defRPr = (para.run as Record<string, unknown> | undefined) ?? {};
      const children = Array.isArray(para.children)
        ? para.children.map((c) =>
            typeof c !== "object" || c === null
              ? { ...defRPr, text: c as string }
              : { ...defRPr, ...(c as object) },
          )
        : undefined;
      const node = ctx.resolveParagraph({
        ...para,
        run: undefined,
        ...(children ? { children } : {}),
      });
      if (node) content.push(node);
    }
  }
  if (content.length === 0) content.push({ type: "paragraph" });
  const { children: _omit, ...geometry } = ws ?? {};
  const node: JSONContent = { type: "wpsShape", content };
  const cleanGeometry = cleanAttrs(geometry as Record<string, unknown>);
  if (Object.keys(cleanGeometry).length > 0) node.attrs = { wpsShape: cleanGeometry };
  return node;
}

// DOCX standalone text-box shape → office-open ParagraphChild `{ wpsShape }`.
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "wpsShape" in child,
  convert: (child, ctx) =>
    resolveWpsShape((child as { wpsShape: Record<string, unknown> }).wpsShape, ctx),
};

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
    const ws = (node.attrs.wpsShape ?? {}) as WpsShapeStandalone;
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

  parseDocxInline,
});
