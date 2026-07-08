import { Extension } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { paragraphMaxRatio, paragraphMaxSizePt } from "../utils/measure";

const key = new PluginKey<DecorationSet>("docenFontMetric");

/** Per-paragraph line-height vars: each paragraph/heading gets two inline CSS
 *  custom properties — --docen-font-metric (CJK-dominant MAX `normal` ratio:
 *  docGrid type=lines is a CJK grid, so the metric follows the CJK run, not a
 *  Latin run alongside it) and --docen-line-base (MAX run size, pt).
 *  lineSpacingToCss resolves line-height as `calc(metric × m × line-base)`, so
 *  the line box scales at the paragraph's dominant font SIZE — matching Word's
 *  line-box rule (a line is as tall as its tallest font), fixing the bug where
 *  a 42pt heading rendered at
 *  the line-height of the inherited 14pt container. ProseMirror APPENDS a node
 *  decoration's `style` to the node's renderHTML style (viewdesc.ts
 *  patchAttributes: `dom.style.cssText +=`), so this coexists with the
 *  paragraph's own line-height/margin inline styles. Recomputed only on
 *  docChanged (caret moves reuse the cached set), so large documents don't
 *  re-probe per keystroke. */
function build(doc: PmNode): DecorationSet {
  const styles = (doc.attrs as { styles?: unknown } | null)?.styles;
  const decos: Decoration[] = [];
  doc.descendants((node, pos, parent) => {
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const ratio = paragraphMaxRatio(node, styles).toFixed(4);
      const size = paragraphMaxSizePt(node, styles);
      const parts = [`--docen-font-metric:${ratio}`, `--docen-line-base:${size}pt`];
      // Table cell: align each line to the grid row (max of the font's natural
      // metric vs the grid pitch) so the row's trHeight atLeast floor — not the
      // line box — governs (Word renders a single-spaced grid row at trHeight).
      // Overrides the paragraph's line-height (ProseMirror appends a node
      // decoration's style, so the later line-height wins); mirrors measure.ts
      // resolveLineHeight(inTable) — the same MAX model as body text (edit == render).
      const parentName = parent?.type.name;
      if (parentName === "tableCell" || parentName === "tableHeader") {
        // Cell line-height = MAX(natural metric, grid pitch). Also size the p to
        // its max run size (--docen-line-base), not the inherited container
        // font-size — a smaller run inside an inherited 12pt cell p leaves a 12pt
        // strut whose baseline alignment makes the line-box ~2px taller than measured.
        parts.push(
          "line-height:calc(max(var(--docen-font-metric) * var(--docen-line-base), var(--docen-line-pitch, 0pt)))",
          "font-size:var(--docen-line-base, 1em)",
        );
      }
      decos.push(Decoration.node(pos, pos + node.nodeSize, { style: parts.join(";") }));
      return false; // don't descend into the textblock's text nodes
    }
    return true;
  });
  return DecorationSet.create(doc, decos);
}

/** Injects --docen-font-metric on every paragraph/heading (the dominant font's
 *  `normal` ratio) so the engine's lineSpacingToCss / sectionLinePitchCss
 *  resolve to the real font metric instead of the 1.2 fallback. Pairs with the
 *  paginator's measure.ts (same paragraphMaxRatio source) for edit == render. */
export const FontMetricDecoration = Extension.create({
  name: "docenFontMetric",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        state: {
          init: (_, state) => build(state.doc),
          apply: (tr, old) => (tr.docChanged ? build(tr.doc) : old),
        },
        props: {
          decorations: (state) => key.getState(state),
        },
      }),
    ];
  },
});
