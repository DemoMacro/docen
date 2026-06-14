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
 * one-liner with no per-node variance, so an extension helper would be pure
 * ceremony.
 */

export const PageBreak = Node.create({
  name: "pageBreak",
  inline: true,
  group: "inline",
  atom: true,

  parseHTML() {
    return [{ tag: 'br[data-type="pageBreak"]' }];
  },

  renderHTML() {
    return ["br", { "data-type": "pageBreak", style: "break-after:page" }];
  },
});
