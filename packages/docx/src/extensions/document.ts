import { Node } from "@tiptap/core";

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
 * None of these render anywhere — they ride the JSON for the converters and
 * the canvas editor's projection to consume.
 *
 * Factory form (`createDocument`): the editor layer may parameterize a
 * different top-level `content` expression but keeps the SAME DOCX attrs.
 * Building it via this factory keeps the Document definition in ONE place
 * (here), instead of `.extend`-overriding this Document and re-stating the
 * attrs. `Document` is the default flat `doc > block+` shape used by the docx
 * package itself.
 */

export function createDocument(content = "block+") {
  return Node.create({
    name: "doc",
    content,
    addAttributes() {
      return {
        styles: attrNative(),
        core: attrNative(),
        sectionProperties: attrNative(),
        sectionHeaders: attrNative(),
        sectionFooters: attrNative(),
        background: attrNative(),
        documentExtras: attrNative(),
        // Bibliography sources (word/bibliography.xml) — the citation dialog's
        // master list and the bibliography block's data.
        bibliography: attrNative(),
        // Source numbering.config (abstractNum definitions) carried verbatim so
        // list markers (glyph/font/indent) round-trip; compile merges it with
        // any regenerated ordered-list definitions.
        numbering: attrNative(),
        // Mail merge state (session-scoped — no OOXML part carries recipient
        // rows): the pasted data source and the Start Mail Merge document
        // kind, consumed by the editor's preview and Finish & Merge.
        recipients: attrNative(),
        mergeType: attrNative(),
      };
    },
  });
}

/** Default flat Document (`doc > block+`) — the DOCX round-trip shape. */
export const Document = createDocument();
