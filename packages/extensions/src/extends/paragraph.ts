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

      // Shading (background color)
      // Maps to CSS background-color and DOCX w:shd
      shading: {
        default: null,
        parseHTML: (element) => {
          const bg = element.style.backgroundColor;
          if (!bg) return null;
          const fill = bg.startsWith("#") ? bg : `#${bg}`;
          return { fill, type: "clear" };
        },
        renderHTML: (attributes) => {
          if (!attributes.shading?.fill) return {};
          return { style: `background-color: ${attributes.shading.fill}` };
        },
      },

      // Borders (not rendered in HTML, only for DOCX)
      borderTop: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
      borderBottom: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
      borderLeft: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
      borderRight: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
    };
  },
});
