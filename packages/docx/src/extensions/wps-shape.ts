import type {
  ParagraphChild,
  ParagraphOptions,
  SectionChild,
  ShapeOptions,
} from "@office-open/docx";
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

/** The standalone text-box shape ParagraphChild branch. */
type WpsBranch = Extract<ParagraphChild, { wpsShape: ShapeOptions }>;

/** Block tag keys — a ShapeTextBoxChild that is not a paragraph carries one of
 *  these (a nested block branch); a ParagraphOptions carries none. */
const BLOCK_TAGS = [
  "paragraph",
  "table",
  "toc",
  "textbox",
  "sdt",
  "altChunk",
  "customXml",
  "bookmarkStart",
  "bookmarkEnd",
  "rawXml",
] as const;

const isBlockBranch = (child: object): child is SectionChild =>
  BLOCK_TAGS.some((tag) => tag in child);

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
function resolveWpsShape(ws: WpsBranch["wpsShape"], ctx: ResolveContext): JSONContent {
  const content: JSONContent[] = [];
  if (ws?.children) {
    for (const child of ws.children) {
      if (typeof child !== "object" || child === null) {
        const node = ctx.resolveParagraph(child);
        if (node) content.push(node);
        continue;
      }
      // A text-box body is paragraphs; a stray nested block branch resolves
      // through the block stream instead of being forced through the
      // paragraph path.
      if (isBlockBranch(child)) {
        const node = ctx.resolveBlock(child);
        if (node) content.push(node);
        continue;
      }
      const para: ParagraphOptions = child;
      // DrawingML defRPr (para.run) is the default run-properties for the box's
      // runs, NOT the OOXML ¶-mark rPr. Merge it into each run (matching the
      // prior atom renderWpsText: {...para.run, ...r}), then drop it from the
      // paragraph (run: undefined): paragraph.ts renders attrs.run.size as
      // ¶-mark line-height, which would override the box's grid line-height —
      // but defRPr is a run default, not a ¶ mark. Round-trip safe — runs carry
      // the full rPr, so compile emits per-run rPr and Word renders identically.
      const defRPr = para.run ?? {};
      const children = Array.isArray(para.children)
        ? para.children.map((c) =>
            typeof c !== "object" || c === null ? { ...defRPr, text: c } : { ...defRPr, ...c },
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
  const cleanGeometry = cleanAttrs(geometry);
  if (Object.keys(cleanGeometry).length > 0) node.attrs = { wpsShape: cleanGeometry };
  return node;
}

// DOCX standalone text-box shape → office-open ParagraphChild `{ wpsShape }`.
export const parseDocxInline: ParseInlineRule<WpsBranch> = {
  match: (child): child is WpsBranch => "wpsShape" in child,
  convert: (child, ctx) => resolveWpsShape(child.wpsShape, ctx),
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
