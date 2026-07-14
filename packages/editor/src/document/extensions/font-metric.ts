import { lineSpacingToCss } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import {
  paragraphHasCjk,
  paragraphMaxRatio,
  paragraphMaxSizePt,
  resolveSpacing,
} from "../utils/measure";

const key = new PluginKey<DecorationSet>("docenFontMetric");

/** Smallest block-level range covering every content change between two docs
 *  — the bounds findDiffStart/findDiffEnd report, expanded to whole top-level
 *  blocks so an incremental decoration rebuild covers the full edited
 *  paragraph(s) without leaving stale spans at the edges. Returns undefined
 *  when the docs have no content difference. Shared with PunctCompression. */
export function getEditedBlocksRange(
  oldDoc: PmNode,
  newDoc: PmNode,
  depth = 1,
): [number, number] | undefined {
  const from = oldDoc.content.findDiffStart(newDoc.content);
  const diffEnd = oldDoc.content.findDiffEnd(newDoc.content);
  if (from === null || !diffEnd) return undefined;
  let start = newDoc.resolve(from);
  if (start.depth < depth) start = newDoc.resolve(Math.min(from + 1, newDoc.nodeSize - 2));
  let end = newDoc.resolve(diffEnd.b);
  if (end.depth < depth) end = newDoc.resolve(Math.max(diffEnd.b - 1, 0));
  return [start.before(depth), end.after(depth)];
}

/** Per-paragraph line-height vars: each paragraph/heading gets two inline CSS
 *  custom properties — --docen-font-metric (CJK-dominant MAX `normal` ratio:
 *  docGrid type=lines is a CJK grid, so the metric follows the CJK run, not a
 *  Latin run alongside it) and --docen-line-base (MAX run size, pt).
 *  lineSpacingToCss resolves line-height as `m × (linePitch | metric×line-base)`, so
 *  the line box scales at the paragraph's dominant font SIZE — matching Word's
 *  line-box rule (a line is as tall as its tallest font), fixing the bug where
 *  a 42pt heading rendered at
 *  the line-height of the inherited 14pt container. ProseMirror APPENDS a node
 *  decoration's `style` to the node's renderHTML style (viewdesc.ts
 *  patchAttributes: `dom.style.cssText +=`), so this coexists with the
 *  paragraph's own line-height/margin inline styles. Recomputed only on
 *  docChanged (caret moves reuse the cached set), so large documents don't
 *  re-probe per keystroke. */
function buildDecos(doc: PmNode, from: number, to: number): Decoration[] {
  const styles = (doc.attrs as { styles?: unknown } | null)?.styles;
  const decos: Decoration[] = [];
  doc.nodesBetween(from, to, (node, pos, parent) => {
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const ratio = paragraphMaxRatio(node, styles).toFixed(4);
      const size = paragraphMaxSizePt(node, styles);
      const parts = [`--docen-font-metric:${ratio}`, `--docen-line-base:${size}pt`];
      const parentName = parent?.type.name;
      const inTable = parentName === "tableCell" || parentName === "tableHeader";
      const attrs = node.attrs as {
        snapToGrid?: boolean | null;
        spacing?: { line?: number | null; lineRule?: string | null } | null;
      };
      if (inTable) {
        // Size the <p> to its max run (--docen-line-base), not the inherited
        // container font-size. line-height is emitted ONLY when the paragraph
        // has no direct spacing.line — renderParagraphStyles reads direct attrs
        // alone, so a style-chain spacing.line (measure's resolveSpacing walks
        // basedOn) needs this decoration to keep the DOM matching measure. A
        // direct spacing.line is already emitted by renderHTML; re-emitting it
        // here made PM's patchAttributes removeProperty('line-height') wipe the
        // renderHTML value whenever this decoration was incrementally rebuilt
        // out (table-split reflow clones cells the decoration set can't map),
        // so cell paragraphs lost their 1.5× and fell back to the 2× grid.
        if (attrs.spacing?.line == null) {
          const cellLh = lineSpacingToCss(resolveSpacing(node, styles));
          parts.push(
            cellLh
              ? `line-height:${cellLh}`
              : "line-height:calc(max(var(--docen-font-metric) * var(--docen-line-base), var(--docen-line-pitch, 0pt)))",
          );
        }
        parts.push("font-size:var(--docen-line-base, 1em)");
      } else if (
        paragraphHasCjk(node, styles) &&
        attrs.snapToGrid !== false &&
        resolveSpacing(node, styles)?.line == null
      ) {
        // CJK-dominant body with NO explicit line spacing: snap UP to a whole
        // pitch multiple (CSS round(up)). docGrid type=lines is a CJK grid — CJK
        // chars align to it (ceil to a whole row); Latin body keeps the section
        // container's MAX from lineSpacingToCss. A paragraph WITH spacing.line
        // (auto/exact/atLeast) is excluded — it overrides the grid and renders at
        // its own line height (multiple×natural / fixed), so no ceil here.
        // pitch=0 (no grid) -> round(up, A, 0)=A, a no-op. Overrides the
        // container's line-height (appended node-deco style).
        parts.push(
          "line-height:calc(round(up, var(--docen-font-metric) * var(--docen-line-base), var(--docen-line-pitch, 0pt)))",
        );
      }
      decos.push(Decoration.node(pos, pos + node.nodeSize, { style: parts.join(";") }));
      return false; // don't descend into the textblock's text nodes
    }
    return true;
  });
  return decos;
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
          init: (_, state) =>
            DecorationSet.create(state.doc, buildDecos(state.doc, 0, state.doc.content.size)),
          apply: (tr, oldSet, oldState, newState) => {
            if (!tr.docChanged) return oldSet;
            // A style-model change (gallery edit, setDoc) rewrites every
            // paragraph's metric — full rebuild. Otherwise rebuild only the
            // edited blocks; the rest map to their new positions.
            const stylesChanged =
              (oldState.doc.attrs as { styles?: unknown }).styles !==
              (newState.doc.attrs as { styles?: unknown }).styles;
            // Pagination reflow splits/merges pages (top-level child count
            // changes). Decoration mapping can't follow a cell cloned across a
            // page split, and the edited-block range may not span every cloned
            // cell — so rebuild fully, or those cells lose their line-height
            // decoration (a style-chain spacing.line that renderHTML can't see)
            // and fall back to the 2× grid.
            const pageCountChanged = oldState.doc.childCount !== newState.doc.childCount;
            if (stylesChanged || pageCountChanged)
              return DecorationSet.create(
                newState.doc,
                buildDecos(newState.doc, 0, newState.doc.content.size),
              );
            const range = getEditedBlocksRange(oldState.doc, newState.doc);
            if (!range) return oldSet.map(tr.mapping, newState.doc);
            let set = oldSet.map(tr.mapping, newState.doc);
            set = set.remove(set.find(range[0], range[1]));
            return set.add(newState.doc, buildDecos(newState.doc, range[0], range[1]));
          },
        },
        props: {
          decorations: (state) => key.getState(state),
        },
      }),
    ];
  },
});
