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
 * one-liner with no per-node variance, so an extension helper would be pure
 * ceremony.
 */

export const ColumnBreak = Node.create({
  name: "columnBreak",
  inline: true,
  group: "inline",
  atom: true,

  parseHTML() {
    return [{ tag: 'br[data-type="columnBreak"]' }];
  },

  renderHTML() {
    return ["br", { "data-type": "columnBreak", style: "break-after:column" }];
  },
});
