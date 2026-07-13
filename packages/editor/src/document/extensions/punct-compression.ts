import { Extension } from "@docen/docx/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

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
 * Mirrors FontMetricDecoration: full rebuild on docChanged, cached set reused
 * for caret/selection moves (no doc change). content-visibility:auto skips
 * layout/paint for off-screen pages, so the decoration spans there cost no
 * layout — only the JS build walk (O(chars)) runs, debounced with the reflow.
 */
const CLOSING = new Set("、。，．：；？！）〉》」』】〕〗〙〛｝");
const OPENING = new Set("（〈《「『【〔〖〘〚｛");

function build(doc: PmNode): DecorationSet {
  const decos: Decoration[] = [];
  doc.nodesBetween(0, doc.content.size, (node, pos) => {
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
  return DecorationSet.create(doc, decos);
}

export const PunctCompressionDecoration = Extension.create({
  name: "docenPunctCompression",
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
