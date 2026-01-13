import { Image as BaseImage } from "@tiptap/extension-image";

/**
 * Custom Image extension based on @tiptap/extension-image
 *
 * Adds rotation support for DOCX round-trip conversion:
 * - Rotation is stored as a number attribute (in degrees)
 * - When rendering to HTML, rotation is converted to CSS transform: rotate()
 * - When parsing from HTML, CSS transform is parsed back to rotation attribute
 * - This enables DOCX import/export to handle rotation while maintaining HTML compatibility
 */
export const Image = BaseImage.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes (src, alt, title, width, height)
      ...this.parent?.(),

      // Add rotation attribute
      rotation: {
        default: null,
        // Parse rotation from CSS transform when importing HTML
        parseHTML: (element) => {
          const style = (element as HTMLElement).getAttribute("style") || "";
          const rotationMatch = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);
          return rotationMatch ? parseFloat(rotationMatch[1]) : null;
        },
        // Render rotation as CSS transform when exporting to HTML
        renderHTML: (attributes) => {
          if (!attributes.rotation) {
            return {};
          }
          return {
            style: `transform: rotate(${attributes.rotation}deg)`,
          };
        },
      },
    };
  },
});
