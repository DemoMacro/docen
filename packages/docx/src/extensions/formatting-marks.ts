import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { Extension } from "../core";

/** Plugin key holding the on/off boolean for formatting marks. Toggled via a
 *  transaction meta so `decorations` recomputes when the user flips the Show/
 *  Hide button. */
const FORMAT_MARKS_KEY = new PluginKey<boolean>("docen-formatting-marks");

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
 */
export const FormattingMarks = Extension.create({
  name: "formattingMarks",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: FORMAT_MARKS_KEY,
        state: {
          init: () => false,
          apply: (tr, value) => {
            const meta = tr.getMeta(FORMAT_MARKS_KEY);
            if (meta === "toggle") return !value;
            return value;
          },
        },
        props: {
          decorations(state) {
            if (!FORMAT_MARKS_KEY.getState(state)) return DecorationSet.empty;
            const decos: Decoration[] = [];
            state.doc.descendants((node, pos) => {
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
            return DecorationSet.create(state.doc, decos);
          },
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
