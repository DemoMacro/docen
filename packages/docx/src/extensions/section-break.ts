import { TextSelection } from "@tiptap/pm/state";

import { Extension } from "../core";

/**
 * SectionBreak — command extension that marks a paragraph as a section boundary.
 *
 * OOXML sections (sectPr) attach to a section's LAST paragraph's pPr, NOT a
 * standalone node. So this extension provides only the `setSectionBreak`
 * command (stamps sectionProperties on the current paragraph); the paragraph
 * extension carries the sectionProperties/sectionHeaders/sectionFooters attrs,
 * and DocxManager splits/merges sections in compile/resolve by reading them off
 * the paragraph.
 *
 * The final section's sectPr rides on doc.attrs.sectionProperties (it lives at
 * <w:body>'s end in OOXML). Single-section documents have no section-carrying
 * paragraph at all.
 *
 * Next Page semantics: `setSectionBreak` stamps sectionProperties on the current
 * paragraph (making it the section's last paragraph) AND inserts a fresh empty
 * paragraph after it (the next section's first paragraph), then moves the
 * selection into that new paragraph. The page-plugin's `forcesPageBreakAfter`
 * treats a sectionProperties-bearing paragraph as a page break, so repaginate
 * pushes the new paragraph onto the next page — and the caret follows. This
 * mirrors Word's "Section Break (Next Page)".
 *
 * Split hygiene: pressing Enter inside a section-carrying paragraph must NOT
 * split it. A split would place the new paragraph past the section boundary —
 * forcesPageBreakAfter then pushes it onto the next page (next section), and
 * splitBlock would copy sectionProperties onto it (a second break mark). Word
 * instead inserts a fresh paragraph BEFORE the section's last paragraph, so the
 * new paragraph stays in this section and the stamped paragraph remains last.
 * The Enter shortcut does exactly that; non-section paragraphs fall through to
 * the default Enter unchanged.
 */
export const SectionBreak = Extension.create({
  name: "sectionBreak",
  // Run before the paragraph extension's own Enter handler so the section
  // split-fix wins; non-section paragraphs return false and fall through.
  priority: 1000,

  addCommands() {
    return {
      // Next Page section break: stamp the current paragraph as its section's
      // last paragraph (forcesPageBreakAfter then closes the page after it),
      // insert a fresh empty paragraph as the next section's first paragraph,
      // and move the caret into it so it lands on the next page after reflow.
      // A new paragraph is inserted rather than split from the current one so
      // it does NOT inherit the just-stamped sectionProperties (which would
      // make it a section boundary too and page-break forever).
      setSectionBreak:
        () =>
        ({ tr, state, dispatch }) => {
          if (!dispatch) return true;
          const { $from } = tr.selection;
          const para = $from.parent;
          if (para.type.name !== "paragraph") return false;
          const paraPos = $from.before($from.depth);
          // 1. Current paragraph becomes its section's last paragraph.
          tr.setNodeMarkup(paraPos, undefined, {
            ...para.attrs,
            sectionProperties: {},
          });
          // 2. Insert a fresh empty paragraph (new section's first paragraph)
          //    and move the caret into it; repaginate pushes it to the next page.
          const paraEnd = paraPos + para.nodeSize;
          tr.insert(paraEnd, state.schema.nodes.paragraph.create());
          tr.setSelection(TextSelection.near(tr.doc.resolve(paraEnd + 1)));
          tr.scrollIntoView();
          return true;
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      Enter: ({ editor }) => {
        const { $from } = editor.state.selection;
        const para = $from.parent;
        // Not a section-carrying paragraph → let the default splitBlock run.
        if (para.attrs.sectionProperties == null) return false;
        // Section-carrying paragraph (its section's last paragraph): insert a
        // fresh empty paragraph BEFORE it, not split it. Splitting would place
        // the new paragraph past the section boundary, where forcesPageBreakAfter
        // shoves it onto the next page (next section); splitBlock would also
        // copy sectionProperties onto it (a second break mark). Inserting before
        // keeps the new paragraph in this section and the stamped paragraph
        // remains the section's last — matching Word.
        const paraPos = $from.before($from.depth);
        return editor
          .chain()
          .insertContentAt(paraPos, { type: para.type.name })
          .setTextSelection(paraPos + 1)
          .run();
      },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    sectionBreak: {
      /** Insert a Next Page section break at the current paragraph. */
      setSectionBreak: () => ReturnType;
    };
  }
}
