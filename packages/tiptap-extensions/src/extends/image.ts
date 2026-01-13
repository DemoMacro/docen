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
      },
    };
  },

  parseHTML() {
    return [
      {
        // Match img tags
        tag: "img[src]",
        // Extract all attributes including rotation from style
        getAttributes: (element: HTMLElement) => {
          const attrs = {
            src: element.getAttribute("src"),
            alt: element.getAttribute("alt"),
            title: element.getAttribute("title"),
            width: element.getAttribute("width"),
            height: element.getAttribute("height"),
          };

          // Parse rotation from CSS transform (only if present)
          const style = element.getAttribute("style") || "";
          const rotationMatch = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);

          // Only add rotation attribute if it exists in HTML
          // Otherwise, TipTap will use the default value (null) from addAttributes()
          return rotationMatch ? { ...attrs, rotation: parseFloat(rotationMatch[1]) } : attrs;
        },
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    // Extract rotation from attributes
    const { rotation, ...otherAttrs } = HTMLAttributes;

    // If rotation exists, add it to style
    if (rotation) {
      const existingStyle = otherAttrs.style || "";
      otherAttrs.style = existingStyle
        ? `${existingStyle}; transform: rotate(${rotation}deg)`
        : `transform: rotate(${rotation}deg)`;
    }

    return ["img", otherAttrs];
  },
});
