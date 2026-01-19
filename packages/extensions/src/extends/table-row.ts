import { TableRow as BaseTableRow } from "../tiptap";

/**
 * Custom TableRow extension with row height support for DOCX round-trip
 *
 * Adds support for row height used in DOCX conversion:
 * - rowHeight: height of the table row (in pixels or CSS units)
 *
 * This attribute maps to CSS height property for HTML rendering
 * and to DOCX w:trPr/w:trHeight for DOCX export/import.
 *
 * Note: Attribute stores CSS value as-is (no unit conversion).
 * Conversion happens in export-docx/import-docx packages.
 */
export const TableRow = BaseTableRow.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes
      ...this.parent?.(),

      // Row height (CSS value: e.g., "20px", "1.5em", "auto")
      // Maps to CSS height and DOCX w:trHeight/@w:val
      rowHeight: {
        default: null,
        parseHTML: (element) => {
          // Try to get height from inline style
          const height = element.style.height;
          if (height) {
            return height;
          }
          // Try to get from height attribute
          const heightAttr = element.getAttribute("height");
          if (heightAttr) {
            return heightAttr;
          }
          return null;
        },
        renderHTML: (attributes) => {
          if (!attributes.rowHeight) {
            return {};
          }
          return {
            style: `height: ${attributes.rowHeight}`,
          };
        },
      },
    };
  },
});
