import { Node } from "../core";

/**
 * ColumnBreak — inline atom node for DOCX column breaks (`<w:br w:type="column"/>`).
 *
 * Sibling to PageBreak and HardBreak: all three render `<w:br>`, differing only
 * by type. Inline (group "inline") to match the run-internal OOXML position,
 * keeping a column break inside its paragraph on round-trip. Visual column
 * layout (paged.js multi-column) is a future concern; the node preserves the
 * break losslessly regardless.
 *
 * The DOCX payload (`{ columnBreak: true }`) is inlined in DocxManager — a
 * one-liner with no per-node variance, so no extension helper for it.
 *
 * `setColumnBreak` only inserts the atom (no paragraph split): a column break
 * does not start a new page, and there is no column layout to reflow yet, so
 * the node is purely for round-trip fidelity until multi-column lands.
 */

export const ColumnBreak = Node.create({
  name: "columnBreak",
  inline: true,
  group: "inline",
  atom: true,

  parseHTML() {
    return [{ tag: 'span[data-type="columnBreak"]' }, { tag: 'br[data-type="columnBreak"]' }];
  },

  renderHTML() {
    // span (not br) so CSS ::before can paint a formatting-marks label.
    return ["span", { "data-type": "columnBreak", style: "break-after:column" }];
  },

  addCommands() {
    return {
      setColumnBreak:
        () =>
        ({ commands }) =>
          commands.insertContent({ type: "columnBreak" }),
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    columnBreak: {
      /** Insert a column break atom at the cursor (round-trip only until
       *  multi-column layout lands). */
      setColumnBreak: () => ReturnType;
    };
  }
}
