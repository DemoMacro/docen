import { Image as BaseImage } from "../tiptap";
import type { SourceRectangleOptions } from "../types";

/**
 * Custom Image extension based on @tiptap/extension-image
 *
 * Adds DOCX-specific attributes for round-trip conversion:
 * - rotation: Image rotation in degrees (rendered as CSS transform)
 * - floating: Image positioning and text wrapping (rendered as CSS float/position/z-index)
 * - outline: Image border/outline options (stored as data-outline attribute)
 * - crop: Image crop rectangle (resized element + object-fit/position)
 *
 * HTML serialization strategy:
 * - rotation: CSS transform merged into floating/crop renderHTML output
 * - floating: CSS float/position/z-index/margin + preserved as data-floating JSON attribute
 * - outline: Preserved as data-outline JSON attribute (no CSS equivalent)
 * - crop: CSS object-fit/position (merged with floating style when both present)
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
      // Note: style output moved to floating/crop renderHTML to avoid TipTap
      // attribute alphabetical merge overriding our CSS
      display: {
        default: null,
        parseHTML: () => (this.options.inline ? "inline-block" : null),
        renderHTML: () => ({}),
      },

      // Add rotation attribute (in degrees)
      rotation: {
        default: null,
        parseHTML: (element) => {
          const style = element.getAttribute("style") || "";
          const rotationMatch = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);
          return rotationMatch ? parseFloat(rotationMatch[1]) : null;
        },
        renderHTML: () => ({}),
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
          const f = attributes.floating;
          const styles: string[] = [];

          // Merge rotation CSS if present (rotation attr defined before floating,
          // its style would be overridden by floating's style in TipTap merge)
          if (attributes.rotation) {
            styles.push(`transform:rotate(${attributes.rotation}deg)`);
          }

          // z-index: behindDocument → below text, otherwise above
          styles.push(f.behindDocument ? "z-index:-1" : "z-index:1");

          // Wrap type → CSS layout
          const wrapType = (f.wrap as { type?: number } | undefined)?.type ?? 0;
          if (wrapType === 0) {
            // NONE: position absolutely
            styles.push("position:absolute");
          } else if (wrapType === 1) {
            // SQUARE: use CSS float
            const align = f.horizontalPosition?.align;
            if (align === "left" || align === "inside") styles.push("float:left");
            else if (align === "right" || align === "outside") styles.push("float:right");
            else styles.push("float:left");
          } else if (wrapType === 2) {
            // TIGHT: float + shape-outside
            styles.push("float:left");
            styles.push("shape-outside:margin-box");
          } else if (wrapType === 3) {
            // TOP_AND_BOTTOM: block-level
            styles.push("display:block");
            styles.push("clear:both");
          } else {
            // Default: inline-block
            styles.push("display:inline-block");
          }

          // Margins (stored in twips, convert to pt for CSS)
          const m = f.margins as { top?: number; bottom?: number; left?: number; right?: number } | undefined;
          if (m) {
            if (m.top) styles.push(`margin-top:${(m.top * 96 / 1440).toFixed(1)}pt`);
            if (m.bottom) styles.push(`margin-bottom:${(m.bottom * 96 / 1440).toFixed(1)}pt`);
            if (m.left) styles.push(`margin-left:${(m.left * 96 / 1440).toFixed(1)}pt`);
            if (m.right) styles.push(`margin-right:${(m.right * 96 / 1440).toFixed(1)}pt`);
          }

          return {
            style: styles.join(";"),
            "data-floating": JSON.stringify(f),
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
