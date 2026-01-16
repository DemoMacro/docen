import { Heading as BaseHeading } from "@tiptap/extension-heading";

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
 * Custom Heading extension with DOCX-compatible style attributes
 *
 * Adds the same paragraph-level formatting as Paragraph extension:
 * - Indentation: left, right, first line
 * - Spacing: before, after
 *
 * This ensures consistency across all block-level elements for DOCX round-trip.
 */
export const Heading = BaseHeading.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes (including level)
      ...this.parent?.(),

      // Left indentation (in pixels)
      indentLeft: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.marginLeft),
        renderHTML: (attributes) =>
          attributes.indentLeft ? { style: `margin-left: ${attributes.indentLeft}px` } : {},
      },

      // Right indentation (in pixels)
      indentRight: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.marginRight),
        renderHTML: (attributes) =>
          attributes.indentRight ? { style: `margin-right: ${attributes.indentRight}px` } : {},
      },

      // First line indentation (in pixels)
      indentFirstLine: {
        default: null,
        parseHTML: (element) => parseIndent(element.style.textIndent),
        renderHTML: (attributes) =>
          attributes.indentFirstLine
            ? { style: `text-indent: ${attributes.indentFirstLine}px` }
            : {},
      },

      // Spacing before heading (in pixels)
      spacingBefore: {
        default: null,
        parseHTML: (element) => parseSpacing(element.style.marginTop),
        renderHTML: (attributes) =>
          attributes.spacingBefore ? { style: `margin-top: ${attributes.spacingBefore}px` } : {},
      },

      // Spacing after heading (in pixels)
      spacingAfter: {
        default: null,
        parseHTML: (element) => parseSpacing(element.style.marginBottom),
        renderHTML: (attributes) =>
          attributes.spacingAfter ? { style: `margin-bottom: ${attributes.spacingAfter}px` } : {},
      },
    };
  },
});
