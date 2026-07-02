import { Document as BaseDocument } from "@tiptap/extension-document";

import { attrNative } from "./utils";

/**
 * Document extension carrying DOCX document-level data through the Tiptap JSON
 * for lossless round-trip (declared as attrs so editor setContent → getJSON
 * preserves them, not just the standalone converters):
 *
 * - `attrs.styles` — office-open `StylesOptions` (styles.xml: importedStyles /
 *   docDefaultsXml / latentStylesXml as raw XML).
 * - `attrs.core` — docProps/core.xml properties (title/creator/description/…,
 *   see DocxCoreProperties in converters/docx.ts).
 * - `attrs.sectionProperties` — the last section's page layout (page size/margin/
 *   orientation, columns, type, grid; intermediate sections carry theirs on
 *   sectionBreak nodes).
 *
 * None rendered to HTML — phase 2 applies styles via injected CSS.
 *
 * Factory form (`createDocument`): the editor layer needs a different top-level
 * content expression (`doc > page+` for the C-route editing schema) but the SAME
 * DOCX attrs. Building it via this factory keeps the Document definition in ONE
 * place (here) — the editor parameterizes only `content`, instead of `.extend`-
 * overriding this Document and re-stating the attrs. `Document` is the default
 * flat `doc > block+` shape used by the docx package itself.
 */

export function createDocument(content = "block+") {
  return BaseDocument.extend({
    content,
    addAttributes() {
      return {
        ...this.parent?.(),

        styles: attrNative(),
        core: attrNative(),
        sectionProperties: attrNative(),
        sectionHeaders: attrNative(),
        sectionFooters: attrNative(),
        background: attrNative(),
        documentExtras: attrNative(),
        // Source numbering.config (abstractNum definitions) carried verbatim so
        // list markers (glyph/font/indent) round-trip; compile merges it with
        // any regenerated ordered-list definitions.
        numbering: attrNative(),
      };
    },
  });
}

/** Default flat Document (`doc > block+`) — the DOCX round-trip shape. */
export const Document = createDocument();
