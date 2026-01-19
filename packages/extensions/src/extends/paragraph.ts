import { Paragraph as BaseParagraph } from "../tiptap";

/**
 * Custom Paragraph extension with DOCX-compatible style attributes
 *
 * Adds support for paragraph-level formatting used in DOCX round-trip conversion:
 * - Indentation: left, right, first line
 * - Spacing: before, after
 *
 * These attributes map to CSS margin properties for HTML rendering
 * and to DOCX paragraph properties for DOCX export/import.
 *
 * Note: Attributes store CSS values as-is (no unit conversion).
 * Conversion happens in export-docx/import-docx packages.
 */
export const Paragraph = BaseParagraph.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes
      ...this.parent?.(),

      // Left indentation (CSS value: e.g., "20px", "1.5rem")
      // Maps to CSS margin-left and DOCX w:ind/@w:left
      indentLeft: {
        default: null,
        parseHTML: (element) => element.style.marginLeft || null,
        renderHTML: (attributes) =>
          attributes.indentLeft ? { style: `margin-left: ${attributes.indentLeft}` } : {},
      },

      // Right indentation (CSS value)
      // Maps to CSS margin-right and DOCX w:ind/@w:right
      indentRight: {
        default: null,
        parseHTML: (element) => element.style.marginRight || null,
        renderHTML: (attributes) =>
          attributes.indentRight ? { style: `margin-right: ${attributes.indentRight}` } : {},
      },

      // First line indentation (CSS value)
      // Maps to CSS text-indent and DOCX w:ind/@w:firstLine
      indentFirstLine: {
        default: null,
        parseHTML: (element) => element.style.textIndent || null,
        renderHTML: (attributes) =>
          attributes.indentFirstLine ? { style: `text-indent: ${attributes.indentFirstLine}` } : {},
      },

      // Spacing before paragraph (CSS value)
      // Maps to CSS margin-top and DOCX w:spacing/@w:before
      spacingBefore: {
        default: null,
        parseHTML: (element) => element.style.marginTop || null,
        renderHTML: (attributes) =>
          attributes.spacingBefore ? { style: `margin-top: ${attributes.spacingBefore}` } : {},
      },

      // Spacing after paragraph (CSS value)
      // Maps to CSS margin-bottom and DOCX w:spacing/@w:after
      spacingAfter: {
        default: null,
        parseHTML: (element) => element.style.marginBottom || null,
        renderHTML: (attributes) =>
          attributes.spacingAfter ? { style: `margin-bottom: ${attributes.spacingAfter}` } : {},
      },
    };
  },
});
