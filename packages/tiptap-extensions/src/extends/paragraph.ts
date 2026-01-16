import { Paragraph as BaseParagraph } from "@tiptap/extension-paragraph";

/**
 * Parse CSS length value to pixels
 * Supports: px, pt, em, rem, %, and unitless values
 */
function parseIndent(value: string | undefined): number | null {
  if (!value) return null;

  // Remove whitespace
  value = value.trim();

  // Match number and optional unit
  const match = value.match(/^([\d.]+)(px|pt|em|rem|%|)?$/);
  if (!match) return null;

  const num = parseFloat(match[1]);
  if (isNaN(num)) return null;

  const unit = match[2] || "px";

  // Convert to pixels (simplified conversion)
  switch (unit) {
    case "px":
      return Math.round(num);
    case "pt":
      return Math.round(num * 1.333); // 1pt = 1.333px
    case "em":
    case "rem":
      return Math.round(num * 16); // Assume 16px base font size
    case "%":
      return Math.round((num * 16) / 100); // % of em, assume 16px base
    default:
      return Math.round(num);
  }
}

/**
 * Parse CSS spacing value to pixels
 * Uses the same logic as parseIndent
 */
function parseSpacing(value: string | undefined): number | null {
  return parseIndent(value);
}

/**
 * Custom Paragraph extension with DOCX-compatible style attributes
 *
 * Adds support for paragraph-level formatting used in DOCX round-trip conversion:
 * - Indentation: left, right, first line
 * - Spacing: before, after
 *
 * These attributes map to CSS margin properties for HTML rendering
 * and to DOCX paragraph properties for DOCX export/import.
 */
export const Paragraph = BaseParagraph.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes
      ...this.parent?.(),

      // Left indentation (in pixels)
      // Maps to CSS margin-left and DOCX w:ind/@w:left
      indentLeft: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.marginLeft),
        renderHTML: (attributes) =>
          attributes.indentLeft ? { style: `margin-left: ${attributes.indentLeft}px` } : {},
      },

      // Right indentation (in pixels)
      // Maps to CSS margin-right and DOCX w:ind/@w:right
      indentRight: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.marginRight),
        renderHTML: (attributes) =>
          attributes.indentRight ? { style: `margin-right: ${attributes.indentRight}px` } : {},
      },

      // First line indentation (in pixels)
      // Maps to CSS text-indent and DOCX w:ind/@w:firstLine
      indentFirstLine: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.textIndent),
        renderHTML: (attributes) =>
          attributes.indentFirstLine
            ? { style: `text-indent: ${attributes.indentFirstLine}px` }
            : {},
      },

      // Spacing before paragraph (in pixels)
      // Maps to CSS margin-top and DOCX w:spacing/@w:before
      spacingBefore: {
        default: null,
        parseHTML: (element) => parseSpacing(element.style.marginTop),
        renderHTML: (attributes) =>
          attributes.spacingBefore ? { style: `margin-top: ${attributes.spacingBefore}px` } : {},
      },

      // Spacing after paragraph (in pixels)
      // Maps to CSS margin-bottom and DOCX w:spacing/@w:after
      spacingAfter: {
        default: null,
        parseHTML: (element) => parseSpacing(element.style.marginBottom),
        renderHTML: (attributes) =>
          attributes.spacingAfter ? { style: `margin-bottom: ${attributes.spacingAfter}px` } : {},
      },
    };
  },
});
