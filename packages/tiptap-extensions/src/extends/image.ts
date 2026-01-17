import { Image as BaseImage } from "@tiptap/extension-image";

/**
 * Custom Image extension based on @tiptap/extension-image
 *
 * Adds DOCX-specific attributes for round-trip conversion:
 * - rotation: Image rotation in degrees (rendered as CSS transform)
 * - floating: Image positioning options (stored as data-floating attribute)
 * - outline: Image border/outline options (stored as data-outline attribute)
 *
 * HTML serialization strategy:
 * - rotation: Mapped to CSS transform: rotate()
 * - floating: Preserved as data-floating JSON attribute (no CSS equivalent)
 * - outline: Preserved as data-outline JSON attribute (no CSS equivalent)
 *
 * Note: floating and outline are DOCX-specific features without direct CSS
 * equivalents. They're preserved in HTML for round-trip conversion but only
 * affect DOCX export/import.
 */
export const Image = BaseImage.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes (src, alt, title, width, height)
      ...this.parent?.(),

      // Add rotation attribute (in degrees)
      rotation: {
        default: null,
        // Parse from CSS transform: rotate(Xdeg)
        parseHTML: (element) => {
          const style = element.getAttribute("style") || "";
          const rotationMatch = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);
          return rotationMatch ? parseFloat(rotationMatch[1]) : null;
        },
        // Render as CSS transform
        renderHTML: (attributes) => {
          if (!attributes.rotation) return {};
          return {
            style: `transform: rotate(${attributes.rotation}deg)`,
          };
        },
      },

      // Add floating attribute for image positioning
      floating: {
        default: null,
        // Parse from data-floating attribute (JSON string)
        parseHTML: (element) => {
          const dataFloating = element.getAttribute("data-floating");
          if (!dataFloating) return null;
          try {
            return JSON.parse(dataFloating);
          } catch {
            return null;
          }
        },
        // Render as data-floating attribute (JSON string)
        renderHTML: (attributes) => {
          if (!attributes.floating) return {};
          return {
            "data-floating": JSON.stringify(attributes.floating),
          };
        },
      },

      // Add outline attribute for image border/outline
      outline: {
        default: null,
        // Parse from data-outline attribute (JSON string)
        parseHTML: (element) => {
          const dataOutline = element.getAttribute("data-outline");
          if (!dataOutline) return null;
          try {
            return JSON.parse(dataOutline);
          } catch {
            return null;
          }
        },
        // Render as data-outline attribute (JSON string)
        renderHTML: (attributes) => {
          if (!attributes.outline) return {};
          return {
            "data-outline": JSON.stringify(attributes.outline),
          };
        },
      },
    };
  },
});
