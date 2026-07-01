import { WpsShape, wpsShapeStyles, type WpsShapeStandalone } from "@docen/docx";

/**
 * Editor-only NodeView for the wpsShape (standalone floating text-box) node.
 *
 * EXTENDS the engine's WpsShape — does not re-create it — so its schema
 * (content:"block+", isolating, parseHTML/renderHTML) is inherited; only the
 * NodeView is added. The engine node is UI-free, so without this the editable
 * body would render via renderHTML (which works) but ProseMirror would not own
 * a contentDOM — the NodeView lets it manage selection, IME, and mutations
 * directly, the same way PageBreakView customizes pageBreak rendering.
 *
 * Two-element structure (mandatory, not stylistic):
 *  - `dom` (outer)        : floating placement + size + rotation + writing-mode.
 *  - `contentDOM` (inner) : box-sizing + fill + outline + textbox insets.
 * rotation/writing-mode live on `dom` ONLY — on the contentDOM they distort the
 * caret rect and break CJK IME composition. wpsShapeStyles (from the engine)
 * computes the split so the editor never re-derives EMU/floating geometry.
 */
export const WpsShapeView = WpsShape.extend({
  addNodeView() {
    return ({ node }) => {
      const ws = (node.attrs.wpsShape ?? {}) as WpsShapeStandalone;
      const { outer, inner, paragraphAnchor } = wpsShapeStyles(ws);
      const dom = document.createElement("div");
      // Mirror renderHTML's data-wps-shape so DOM/CSS selectors (and any DOM
      // reader) treat the NodeView identically to the rendered node.
      dom.setAttribute("data-wps-shape", JSON.stringify(ws));
      dom.style.cssText = outer;
      if (paragraphAnchor) dom.setAttribute("data-float-anchor", "paragraph");
      const contentDOM = document.createElement("div");
      contentDOM.style.cssText = inner;
      dom.appendChild(contentDOM);
      return { dom, contentDOM };
    };
  },
});
