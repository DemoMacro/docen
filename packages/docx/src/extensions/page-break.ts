import { Plugin, TextSelection } from "@tiptap/pm/state";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";

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

// DOCX `<w:br w:type="page"/>` → office-open ParagraphChild `{ pageBreak: true }`.
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "pageBreak" in child,
  convert: () => ({ type: "pageBreak" }),
};

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

  parseDocxInline,

  addProseMirrorPlugins() {
    // The C-route paginator forces a new page after a paragraph CONTAINING a
    // pageBreak, and a single paragraph cannot span pages. So anything that
    // lands AFTER a pageBreak in the SAME paragraph (typed, pasted, or dragged
    // there — the atom is inline, so the spot right after it is editable) would
    // render beside the break instead of flowing to the next page. Word treats
    // a page break as a paragraph terminator; mirror that by splitting the
    // paragraph right after any pageBreak that still has trailing content. The
    // trailing content becomes a fresh paragraph the paginator moves on.
    return [
      new Plugin({
        appendTransaction: (transactions, _oldState, newState) => {
          if (!transactions.some((tr) => tr.docChanged)) return null;
          const doc = newState.doc;
          const tr = newState.tr;
          const splits: number[] = [];
          doc.descendants((node, pos) => {
            if (!node.isTextblock) return;
            const last = node.childCount - 1;
            node.forEach((child, offset, index) => {
              if (child.type.name === "pageBreak" && index < last) {
                // `pos` is the paragraph's start (before <p>); content begins at
                // pos+1, so the doc position right after this pageBreak atom is
                // pos + 1 + offset + nodeSize.
                splits.push(pos + 1 + offset + child.nodeSize);
              }
            });
            return false;
          });
          if (splits.length === 0) return null;
          // Split from the end backward so earlier positions stay valid.
          splits.sort((a, b) => b - a);
          for (const splitPos of splits) tr.split(splitPos, 1);
          tr.setSelection(newState.selection.map(tr.doc, tr.mapping));
          return tr;
        },
      }),
    ];
  },

  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ tr, state, dispatch }) => {
          if (!dispatch) return true;
          const { $from } = state.selection;
          // Back-to-back break: the caret sits right after an existing pageBreak
          // atom. Word fires each <w:br type=page> as its own page break (two in
          // a row leave a blank page between). The C-route paginator breaks at
          // the paragraph level, so a second atom stacked in the SAME paragraph
          // would share one break and both land on one page. Start a fresh break
          // paragraph after this one so each atom owns its own page.
          const depth = $from.depth;
          const prev = $from.index(depth) > 0 ? $from.parent.child($from.index(depth) - 1) : null;
          if (prev?.type.name === "pageBreak") {
            const after = $from.after(depth);
            const carrier = state.schema.nodes.paragraph.create(null, [
              state.schema.nodes.pageBreak.create(),
            ]);
            tr.insert(after, carrier);
            tr.setSelection(TextSelection.near(tr.doc.resolve(after + carrier.nodeSize)));
            tr.scrollIntoView();
            return true;
          }
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
