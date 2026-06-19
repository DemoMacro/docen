import { Document as BaseDocument } from "./tiptap";

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
 */
const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

export const Document = BaseDocument.extend({
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
    };
  },
});
