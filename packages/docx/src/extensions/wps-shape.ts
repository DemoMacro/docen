import { convertEmuToPixels } from "@office-open/core";
import type { DOMOutputSpec } from "@tiptap/pm/model";

import { Node } from "../core";
import { floatAnchorScope, floatingToStyles } from "./utils";
import { renderWpsInterior, type WpsData } from "./wpg-group";

/**
 * wpsShape — inline atom carrying a standalone DOCX text-box shape
 * (wp:anchor > wps:wsp > wps:txbx; NOT inside a wpg group) as an opaque blob.
 * Mirrors the office-open `WpsShapeRunOptions` / ParagraphChild `wpsShape` field
 * verbatim in attrs.wpsShape. Unlike a group's interior wps children (laid out in
 * the group's coordinate space), this shape floats on its own anchor — placement
 * comes from `floating` (rendered via floatingToStyles), while its body
 * (fill/outline/insets/text) reuses the group's renderWpsInterior so the two
 * renderers stay identical.
 */

/** wpsShape data (WpsShapeRunOptions subset used for rendering). */
interface WpsShapeData extends WpsData {
  transformation?: { width?: number; height?: number };
  floating?: unknown;
}

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
  atom: true,

  addAttributes() {
    return {
      wpsShape: attrWpsShape(),
    };
  },

  parseHTML() {
    return [{ tag: "div[data-wps-shape]" }];
  },

  renderHTML({
    node,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const ws = (node.attrs.wpsShape as WpsShapeData | null) ?? {};
    // office-open 0.10.4+ parses extent as EMU verbatim (was pixels); convert to px.
    const w = ws.transformation?.width != null ? convertEmuToPixels(ws.transformation.width) : 0;
    const h = ws.transformation?.height != null ? convertEmuToPixels(ws.transformation.height) : 0;
    // box-sizing:border-box so the shape's width/height is its outer box
    // (matching Word's extent), not content-box (which adds padding on top).
    const sizeStyle = `width:${w}px;height:${h}px;box-sizing:border-box`;
    if (ws.floating) {
      // A floating text box (wp:anchor wrapNone) overlays its anchor paragraph
      // instead of claiming a line in the flow — same anchor CSS as images and
      // wpg groups (position:absolute at the EMU offset). A paragraph-anchored
      // box (vRelative "paragraph") resolves its top/left from the anchor <p>
      // (data-float-anchor → editor CSS makes the <p> relative); otherwise it
      // anchors to the page box and floats over the body.
      const pos = [
        ...floatingToStyles(ws.floating, undefined, ws.transformation?.width),
        sizeStyle,
      ].join(";");
      const attrs: Record<string, string> = { "data-wps-shape": "" };
      if (floatAnchorScope(ws.floating) === "paragraph") {
        attrs["data-float-anchor"] = "paragraph";
      }
      return renderWpsInterior(ws, pos, { attrs }) as unknown as DOMOutputSpec;
    }
    // An inline wps (rare — no wp:anchor) flows with the text as an inline block.
    return renderWpsInterior(ws, `display:inline-block;vertical-align:middle;${sizeStyle}`, {
      attrs: { "data-wps-shape": "" },
    }) as unknown as DOMOutputSpec;
  },
});
