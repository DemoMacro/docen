import { BulletList as BulletListBase } from "./tiptap";

/**
 * BulletList — carries the source DOCX abstractNum reference (when the list came
 * from parseDOCX) so the round-trip reuses the original numbering definition
 * (custom bullet glyph/font/indent, e.g. a Wingdings marker) instead of
 * regenerating the default. The `numbering` attr is DOCX-only (not rendered
 * to HTML); lists created in the editor carry null and compile to the default
 * bullet.
 */
export const BulletList = BulletListBase.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      numbering: { default: null, rendered: false },
    };
  },
});
