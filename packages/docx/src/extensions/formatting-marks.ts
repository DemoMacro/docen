import type { Node } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { Extension } from "../core";

type FormattingMarksState = { enabled: boolean; decos: DecorationSet };

/** Plugin key holding {enabled, decos} for the formatting-marks plugin.
 *  `enabled` is toggled by a transaction meta; `decos` caches the widget set so
 *  a pure selection/cursor transaction (doc unchanged) reuses it instead of
 *  re-traversing every textblock. The un-cached version ran that traversal on
 *  every selectionUpdate — 200ms+ on a 1000-page document — making the caret
 *  and selection feel frozen. */
const FORMAT_MARKS_KEY = new PluginKey<FormattingMarksState>("docen-formatting-marks");

/** Build a widget set with one paragraph-mark widget at each textblock's
 *  closing position. Called only on doc change or marks toggle. */
function buildParagraphMarkDecos(doc: Node): DecorationSet {
  const decos: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return;
    decos.push(
      Decoration.widget(
        pos + node.nodeSize - 1,
        () => {
          const span = document.createElement("span");
          span.className = "docen-para-mark";
          span.contentEditable = "false";
          span.textContent = "↲";
          return span;
        },
        { side: 1 },
      ),
    );
  });
  return DecorationSet.create(doc, decos);
}

/**
 * FormattingMarks — paints the non-printing paragraph mark via ProseMirror
 * widget decorations.
 *
 * CSS `p::after { content: "…" }` does NOT work on a ProseMirror-managed
 * paragraph: the view owns the DOM and the trailing-break kludge pushes the
 * pseudo-element off the visible line, so the mark never appears (verified).
 * Marijn Haverbeke endorses widget decorations for exactly this use case
 * (https://discuss.prosemirror.net/t/1442). Each textblock gets a widget at its
 * closing position (inside the textblock, so it stays on the content's line)
 * holding a non-editable `<span class="docen-para-mark">`. `side: 1` keeps the
 * mark to the right of a cursor resting at the paragraph end (the cursor always
 * precedes its own paragraph mark).
 *
 * Glyph: a down-then-left return arrow (↲, U+21B2 "downwards arrow with tip
 * leftwards" — the Enter-key glyph). Word/WPS use the pilcrow ¶ for paragraph
 * marks; that down-then-left arrow is Word's line-break glyph — adopted here as
 * the paragraph mark per the project's visual preference.
 *
 * Performance: the widget set is cached in plugin state and rebuilt only on a
 * doc change or a marks toggle — NOT on every selectionUpdate. Without the
 * cache, each caret move / selection re-traverses the whole document (one
 * widget per textblock) and rebuilds the set, which costs 200ms+ on 1000-page
 * documents. Pure selection transactions now reuse the cached set in O(1).
 */
export const FormattingMarks = Extension.create({
  name: "formattingMarks",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: FORMAT_MARKS_KEY,
        state: {
          init: () => ({ enabled: false, decos: DecorationSet.empty }),
          apply: (tr, prev) => {
            const toggled = tr.getMeta(FORMAT_MARKS_KEY) === "toggle";
            const enabled = toggled ? !prev.enabled : prev.enabled;
            // Marks off → empty set. Marks on + (toggled | docChanged) →
            // rebuild from tr.doc. Pure selection transactions reuse the
            // cached set — the caret/selection hot path.
            if (!enabled) return { enabled: false, decos: DecorationSet.empty };
            if (toggled || tr.docChanged)
              return { enabled, decos: buildParagraphMarkDecos(tr.doc) };
            return { enabled, decos: prev.decos };
          },
        },
        props: {
          decorations: (state) => FORMAT_MARKS_KEY.getState(state)!.decos,
        },
      }),
    ];
  },

  addCommands() {
    return {
      toggleFormattingMarks:
        () =>
        ({ tr, dispatch }) => {
          if (dispatch) dispatch(tr.setMeta(FORMAT_MARKS_KEY, "toggle"));
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    formattingMarks: {
      /** Toggle the non-printing paragraph marks on or off. */
      toggleFormattingMarks: () => ReturnType;
    };
  }
}
