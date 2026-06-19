import { Document as BaseDocument } from "./tiptap";

/**
 * Document extension carrying the DOCX styles library through the Tiptap JSON
 * for lossless round-trip.
 *
 * `attrs.styles` mirrors office-open's `StylesOptions`. `resolveDocument`
 * stores the parsed styles here; `compileDocument` reads it back, so the
 * styles.xml part round-trips verbatim (`importedStyles` / `docDefaultsXml` /
 * `latentStylesXml` are raw XML). Not rendered to HTML — phase 2 applies
 * styles via injected CSS.
 */
const attrNative = () => ({ default: null, parseHTML: () => null, rendered: false });

export const Document = BaseDocument.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      styles: attrNative(),
    };
  },
});
