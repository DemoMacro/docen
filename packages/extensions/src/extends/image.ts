import { Image as BaseImage } from "../tiptap";
import type { SourceRectangleOptions } from "../types";

/**
 * Custom Image extension based on @tiptap/extension-image
 *
 * Adds DOCX-specific attributes for round-trip conversion:
 * - rotation: Image rotation in degrees (rendered as CSS transform)
 * - floating: Image positioning options (stored as data-floating attribute)
 * - outline: Image border/outline options (stored as data-outline attribute)
 * - crop: Image crop rectangle (resized element + object-fit/position)
 *
 * HTML serialization strategy:
 * - rotation: Mapped to CSS transform: rotate()
 * - floating: Preserved as data-floating JSON attribute (no CSS equivalent)
 * - outline: Preserved as data-outline JSON attribute (no CSS equivalent)
 * - crop: object-fit: cover + object-position (element size already reflects crop)
 */

/**
 * Generate crop rendering attributes.
 * Uses object-fit/position to show the correct portion within the element's
 * existing dimensions (which already reflect the cropped bounding box from wp:extent).
 */
export function renderCropAttrs(crop: SourceRectangleOptions): { style: string } {
  const leftPct = (crop.left || 0) / 100000;
  const topPct = (crop.top || 0) / 100000;
  const rightPct = (crop.right || 0) / 100000;
  const bottomPct = (crop.bottom || 0) / 100000;

  const visibleWidthPct = 1 - leftPct - rightPct;
  const visibleHeightPct = 1 - topPct - bottomPct;

  const posX = visibleWidthPct > 0 ? (leftPct / visibleWidthPct) * 100 : 0;
  const posY = visibleHeightPct > 0 ? (topPct / visibleHeightPct) * 100 : 0;

  return {
    style: `object-fit:cover;object-position:${posX.toFixed(2)}% ${posY.toFixed(2)}%`,
  };
}

export const Image = BaseImage.extend({
  addAttributes() {
    return {
      // Inherit all parent attributes (src, alt, title, width, height)
      ...this.parent?.(),

      // Set display to inline-block for inline images to make text-align work properly
      display: {
        default: null,
        parseHTML: () => (this.options.inline ? "inline-block" : null),
        renderHTML: () => (this.options.inline ? { style: "display: inline-block" } : {}),
      },

      // Add rotation attribute (in degrees)
      rotation: {
        default: null,
        parseHTML: (element) => {
          const style = element.getAttribute("style") || "";
          const rotationMatch = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);
          return rotationMatch ? parseFloat(rotationMatch[1]) : null;
        },
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
        parseHTML: (element) => {
          const dataFloating = element.getAttribute("data-floating");
          if (!dataFloating) return null;
          try {
            return JSON.parse(dataFloating);
          } catch {
            return null;
          }
        },
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
        parseHTML: (element) => {
          const dataOutline = element.getAttribute("data-outline");
          if (!dataOutline) return null;
          try {
            return JSON.parse(dataOutline);
          } catch {
            return null;
          }
        },
        renderHTML: (attributes) => {
          if (!attributes.outline) return {};
          return {
            "data-outline": JSON.stringify(attributes.outline),
          };
        },
      },

      // Add crop attribute for image cropping (DOCX srcRect values)
      crop: {
        default: null,
        parseHTML: (element) => {
          const dataCrop = element.getAttribute("data-crop");
          if (!dataCrop) return null;
          try {
            return JSON.parse(dataCrop);
          } catch {
            return null;
          }
        },
        // object-fit/position for correct portion (element size already reflects crop)
        renderHTML: (attributes) => {
          if (!attributes.crop) return {};
          const cropAttrs = renderCropAttrs(attributes.crop);
          return {
            ...cropAttrs,
            "data-crop": JSON.stringify(attributes.crop),
          };
        },
      },
    };
  },
});
