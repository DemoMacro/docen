import type { JSONContent } from "@tiptap/core";

import { Image as BaseImage } from "./tiptap";

type CropRect = { left?: number; top?: number; right?: number; bottom?: number };

/**
 * Custom Image extension with node-level renderHTML + renderDocx/parseDocx.
 *
 * Attrs:
 *  - src/alt/title/width/height: Tiptap structural names (kept verbatim so base
 *    image commands work).
 *  - rotation: editor display only (CSS transform) but also carried through DOCX
 *    via transformation.rotation (MediaTransformation.rotation).
 *  - floating/outline: nested office-open objects (Floating / OutlineOptions).
 *  - crop: nested office-open SourceRectangleOptions (srcRect).
 *  - display: editor-only display hint, no OOXML equivalent.
 *
 * DOCX round-trip is near-identity: renderDocx packs attrs into CoreImageOptions;
 * parseDocx unpacks them back. src is a data URL ↔ { type, data } base64.
 * Node-level renderHTML solves the style merge problem (rotation + floating).
 */

// ── DOCX serialization (module-level, exported for DocxManager) ──

/**
 * Tiptap JSON image node → CoreImageOptions-shaped object.
 *
 * Returns `{ image: ImageOptions }` (structural wrapper) or null when no
 * embedded image data is available (external URLs need pre-fetching).
 * rotation is carried via transformation.rotation (not dropped).
 */
export function renderDocx(node: JSONContent): Record<string, unknown> | null {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const imageOpts: Record<string, unknown> = {};

  // src (data URL) → { type, data } base64-decoded bytes
  const src = attrs.src as string | undefined;
  if (src?.startsWith("data:image/")) {
    const match = src.match(/^data:image\/([\w.+-]+);base64,(.+)$/);
    if (match) {
      imageOpts.type = match[1] === "jpeg" ? "jpg" : match[1];
      const binary = atob(match[2]);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      imageOpts.data = bytes;
    }
  }

  // Cannot generate an image run without embedded data (external URLs need pre-fetching)
  if (!imageOpts.data) return null;

  // transformation: width/height are required by OOXML MediaTransformation —
  // default when absent (editor/prepare step normally supplies real dimensions).
  // rotation is an optional editor attr carried via transformation.rotation.
  const width = (attrs.width as number | undefined) ?? 600;
  const height = (attrs.height as number | undefined) ?? 400;
  const transformation: Record<string, unknown> = { width, height };
  const rotation = attrs.rotation as number | undefined;
  if (rotation != null) transformation.rotation = rotation;
  imageOpts.transformation = transformation;

  // altText: alt → name, title → description (DocPropertiesOptions)
  const altText: Record<string, string> = {};
  if (attrs.alt) altText.name = attrs.alt as string;
  if (attrs.title) altText.description = attrs.title as string;
  if (Object.keys(altText).length > 0) imageOpts.altText = altText;

  // Near-identity pass-through for nested office-open objects
  if (attrs.floating) imageOpts.floating = attrs.floating;
  if (attrs.crop) imageOpts.srcRect = attrs.crop;
  if (attrs.outline) imageOpts.outline = attrs.outline;

  return { image: imageOpts };
}

/**
 * ImageOptions-shaped object → Tiptap attrs.
 *
 * Near-identity unpack: transformation → width/height/rotation, altText → alt/title,
 * floating/srcRect(→crop)/outline passed through verbatim. src is reconstructed by
 * DocxManager from the image data bytes (kept out of parseDocx).
 */
export function parseDocx(imageOpts: Record<string, unknown>): Record<string, unknown> {
  const opts = imageOpts ?? {};
  const attrs: Record<string, unknown> = {};

  // transformation → width/height/rotation (structural Tiptap attrs)
  const transformation = opts.transformation as Record<string, unknown> | undefined;
  if (transformation) {
    if (typeof transformation.width === "number") attrs.width = transformation.width;
    if (typeof transformation.height === "number") attrs.height = transformation.height;
    if (typeof transformation.rotation === "number") attrs.rotation = transformation.rotation;
  }

  // altText → alt/title
  const altText = opts.altText as Record<string, unknown> | undefined;
  if (altText) {
    if (altText.name) attrs.alt = altText.name;
    if (altText.description) attrs.title = altText.description;
  }

  // Near-identity pass-through for nested office-open objects
  if (opts.floating) attrs.floating = opts.floating;
  if (opts.srcRect) attrs.crop = opts.srcRect;
  if (opts.outline) attrs.outline = opts.outline;

  return attrs;
}

// ── Node-level renderHTML helpers ──

export function renderCropAttrs(crop: Record<string, unknown> | CropRect): { style: string } {
  const c = crop as CropRect;
  const leftPct = (c.left || 0) / 100000;
  const topPct = (c.top || 0) / 100000;
  const rightPct = (c.right || 0) / 100000;
  const bottomPct = (c.bottom || 0) / 100000;

  const visibleWidthPct = 1 - leftPct - rightPct;
  const visibleHeightPct = 1 - topPct - bottomPct;

  const posX = visibleWidthPct > 0 ? (leftPct / visibleWidthPct) * 100 : 0;
  const posY = visibleHeightPct > 0 ? (topPct / visibleHeightPct) * 100 : 0;

  return {
    style: `object-fit:cover;object-position:${posX.toFixed(2)}% ${posY.toFixed(2)}%`,
  };
}

function renderImageStyles(attrs: Record<string, unknown>): string[] {
  const styles: string[] = [];

  if (attrs.display) styles.push(`display:${attrs.display as string}`);

  if (attrs.rotation) {
    styles.push(`transform:rotate(${attrs.rotation as number}deg)`);
  }

  const f = attrs.floating as Record<string, unknown> | null;
  if (f) {
    styles.push(f.behindDocument ? "z-index:-1" : "z-index:1");

    const wrapType = ((f.wrap as Record<string, unknown> | undefined)?.type as number) ?? 0;
    if (wrapType === 0) {
      styles.push("position:absolute");
    } else if (wrapType === 1) {
      const align = (f.horizontalPosition as Record<string, unknown> | undefined)?.align;
      if (align === "left" || align === "inside") styles.push("float:left");
      else if (align === "right" || align === "outside") styles.push("float:right");
      else styles.push("float:left");
    } else if (wrapType === 2) {
      styles.push("float:left");
      styles.push("shape-outside:margin-box");
    } else if (wrapType === 3) {
      styles.push("display:block");
      styles.push("clear:both");
    } else {
      styles.push("display:inline-block");
    }

    const m = f.margins as
      | { top?: number; bottom?: number; left?: number; right?: number }
      | undefined;
    if (m) {
      if (m.top) styles.push(`margin-top:${((m.top * 96) / 1440).toFixed(1)}pt`);
      if (m.bottom) styles.push(`margin-bottom:${((m.bottom * 96) / 1440).toFixed(1)}pt`);
      if (m.left) styles.push(`margin-left:${((m.left * 96) / 1440).toFixed(1)}pt`);
      if (m.right) styles.push(`margin-right:${((m.right * 96) / 1440).toFixed(1)}pt`);
    }
  }

  if (attrs.crop) {
    const cropAttrs = renderCropAttrs(attrs.crop as CropRect);
    styles.push(cropAttrs.style);
  }

  return styles;
}

// ── Extension ──

export const Image = BaseImage.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      // Editor-only display hint (no OOXML equivalent; not round-tripped)
      display: {
        default: null,
        rendered: false,
        parseHTML: () => (this.options.inline ? "inline-block" : null),
      },

      // Editor display + DOCX transformation.rotation (degrees)
      rotation: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const style = element.getAttribute("style") || "";
          const match = style.match(/transform:\s*rotate\(([\d.]+)deg\)/);
          return match ? parseFloat(match[1]) : null;
        },
      },

      // Nested office-open Floating (JSON in data-floating; CSS rendered)
      floating: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute("data-floating");
          if (!raw) return null;
          try {
            return JSON.parse(raw);
          } catch {
            return null;
          }
        },
      },

      // Nested office-open OutlineOptions (JSON in data-outline)
      outline: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute("data-outline");
          if (!raw) return null;
          try {
            return JSON.parse(raw);
          } catch {
            return null;
          }
        },
      },

      // Nested office-open SourceRectangleOptions (JSON in data-crop)
      crop: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute("data-crop");
          if (!raw) return null;
          try {
            return JSON.parse(raw);
          } catch {
            return null;
          }
        },
      },
    };
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const styles = renderImageStyles(node.attrs);
    const attrs: Record<string, unknown> = { ...HTMLAttributes };

    if (styles.length > 0) attrs.style = styles.join(";");
    if (node.attrs.floating) attrs["data-floating"] = JSON.stringify(node.attrs.floating);
    if (node.attrs.outline) attrs["data-outline"] = JSON.stringify(node.attrs.outline);
    if (node.attrs.crop) attrs["data-crop"] = JSON.stringify(node.attrs.crop);

    return ["img", attrs] as const;
  },

  renderDocx: renderDocx as (node: JSONContent) => Record<string, unknown>,
  parseDocx: parseDocx as (opts: Record<string, unknown>) => Record<string, unknown>,
});
