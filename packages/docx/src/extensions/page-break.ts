import { Node } from "../core";

/**
 * PageBreak — inline atom node for DOCX page breaks (`<w:br w:type="page"/>`).
 *
 * OOXML page breaks live inside a run (inline), so the Tiptap node is inline
 * too (group "inline"). This keeps a break inside its paragraph on round-trip
 * — a break mid-paragraph stays mid-paragraph instead of splitting the
 * paragraph into separate blocks (which would lose the original structure).
 * Sibling to HardBreak: both render `<w:br>`, differing only by type.
 *
 * The DOCX payload (`{ pageBreak: true }`) is inlined in DocxManager — a
 * one-liner with no per-node variance, so no extension helper for it.
 *
 * Inserting a page break must also visually reflow (the C-route paginator only
 * forces a new page after a block *containing* a pageBreak atom). So
 * `setPageBreak` inserts the atom then splits the paragraph — the atom lands at
 * the end of the first half, the trailing text becomes a fresh paragraph, and
 * the paginator's `forcesPageBreakAfter` moves it to the next page (matching
 * Word's Ctrl+Enter behavior).
 */

export const PageBreak = Node.create({
  name: "pageBreak",
  inline: true,
  group: "inline",
  atom: true,

  parseHTML() {
    // span is the current render; br kept for legacy HTML round-trip.
    return [{ tag: 'span[data-type="pageBreak"]' }, { tag: 'br[data-type="pageBreak"]' }];
  },

  renderHTML() {
    // span (not br) so CSS ::before can paint the dashed "Page Break" rule
    // shown by the show-marks formatting-marks toggle (br is void — no
    // pseudo-element). The DOCX payload stays <w:br w:type="page"/>.
    return ["span", { "data-type": "pageBreak", style: "break-after:page" }];
  },

  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ tr, state, dispatch }) => {
          if (!dispatch) return true;
          // chain().insertContent().splitBlock() is unreliable — splitBlock
          // misses paragraph tails (empty para / end-of-doc), so the break
          // never reflows. Operating on the raw tr splits at the exact
          // post-insert position so the trailing content always starts a
          // new page (Word's Ctrl+Enter behavior).
          tr.replaceSelectionWith(state.schema.nodes.pageBreak.create());
          tr.split(tr.selection.from, 1);
          tr.scrollIntoView();
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    pageBreak: {
      /** Insert a page break at the cursor and split the paragraph so the
       *  trailing content reflows to the next page. */
      setPageBreak: () => ReturnType;
    };
  }
}
