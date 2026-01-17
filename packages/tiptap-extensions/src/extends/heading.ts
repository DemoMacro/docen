import { Heading as BaseHeading } from "@tiptap/extension-heading";

/**
 * Custom Heading extension with DOCX-compatible style attributes
 *
 * Adds the same paragraph-level formatting as Paragraph extension:
 * - Indentation: left, right, first line
 * - Spacing: before, after
 *
 * This ensures consistency across all block-level elements for DOCX round-trip.
 *
 * Note: Attributes store CSS values as-is (no unit conversion).
 * Conversion happens in export-docx/import-docx packages.
 */
export const Heading = BaseHeading.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes (including level)
      ...this.parent?.(),

      // Left indentation (CSS value: e.g., "20px", "1.5rem")
      indentLeft: {
        default: null,
        parseHTML: (element) => element.style.marginLeft || null,
        renderHTML: (attributes) =>
          attributes.indentLeft ? { style: `margin-left: ${attributes.indentLeft}` } : {},
      },

      // Right indentation (CSS value)
      indentRight: {
        default: null,
        parseHTML: (element) => element.style.marginRight || null,
        renderHTML: (attributes) =>
          attributes.indentRight ? { style: `margin-right: ${attributes.indentRight}` } : {},
      },

      // First line indentation (CSS value)
      indentFirstLine: {
        default: null,
        parseHTML: (element) => element.style.textIndent || null,
        renderHTML: (attributes) =>
          attributes.indentFirstLine ? { style: `text-indent: ${attributes.indentFirstLine}` } : {},
      },

      // Spacing before heading (CSS value)
      spacingBefore: {
        default: null,
        parseHTML: (element) => element.style.marginTop || null,
        renderHTML: (attributes) =>
          attributes.spacingBefore ? { style: `margin-top: ${attributes.spacingBefore}` } : {},
      },

      // Spacing after heading (CSS value)
      spacingAfter: {
        default: null,
        parseHTML: (element) => element.style.marginBottom || null,
        renderHTML: (attributes) =>
          attributes.spacingAfter ? { style: `margin-bottom: ${attributes.spacingAfter}` } : {},
      },
    };
  },
});
