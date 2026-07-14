import { Extension } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { getEditedBlocksRange } from "./font-metric";

const key = new PluginKey<DecorationSet>("docenPunctCompression");

/**
 * Fullwidth CJK punctuation carries a half-em side bearing that Word/CJK
 * typography removes to tighten lines. The Pretext measure patch
 * (node_modules/@chenglou/pretext/dist/measurement.js, `punctTrimCorrection`)
 * subtracts fontSize/2 per compressible mark; this decoration applies the
 * matching negative-margin trim in the DOM render so measure == render. The two
 * sets below MUST stay in lock-step with the patch's fullwidthClosingPunct /
 * fullwidthOpeningPunct — closing marks sit ink in the lower-left of the em box
 * (trailing/right bearing removed → margin-inline-end), opening marks sit ink on
 * the right (leading/left bearing removed → margin-inline-start). We never
 * enable CSS text-spacing-trim, so fonts whose OpenType halt/chws features would
 * otherwise do this stay on the same one geometric rule regardless of font.
 *
 * Mirrors FontMetricDecoration: incremental rebuild on docChanged — only the
 * edited blocks are rescanned, the rest map to their new positions; the cached
 * set is reused for caret/selection moves (no doc change). content-visibility:
 * auto skips layout/paint for off-screen pages, so the decoration spans there
 * cost no layout — only the JS build walk (now O(edited chars) instead of
 * O(whole doc)) runs, debounced with the reflow.
 */
const CLOSING = new Set("、。，．：；？！）〉》」』】〕〗〛｝");
const OPENING = new Set("（〈《「『【〔〖〘〚｛");

function buildDecos(doc: PmNode, from: number, to: number): Decoration[] {
  const decos: Decoration[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText && node.text) {
      const text = node.text;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const cls = CLOSING.has(ch)
          ? "docen-punct-close"
          : OPENING.has(ch)
            ? "docen-punct-open"
            : null;
        if (cls !== null) decos.push(Decoration.inline(pos + i, pos + i + 1, { class: cls }));
      }
    }
    return true;
  });
  return decos;
}

export const PunctCompressionDecoration = Extension.create({
  name: "docenPunctCompression",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        state: {
          init: (_, state) =>
            DecorationSet.create(state.doc, buildDecos(state.doc, 0, state.doc.content.size)),
          apply: (tr, oldSet, oldState, newState) => {
            if (!tr.docChanged) return oldSet;
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
