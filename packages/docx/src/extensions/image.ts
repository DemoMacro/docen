import type { JSONContent } from "../core";
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

/** Attribute spec for a nested office-open value stored as JSON in a data-* attr. */
const attrDataJson = (name: string) => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute(name);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

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
  // 0.9.7+ fidelity fields (office-open parses + stringifies each verbatim)
  if (attrs.nonVisualProperties) imageOpts.nonVisualProperties = attrs.nonVisualProperties;
  if (attrs.effectExtent) transformation.effectExtent = attrs.effectExtent;
  if (attrs.graphicFrameLocks) imageOpts.graphicFrameLocks = attrs.graphicFrameLocks;
  if (attrs.blipEffects) imageOpts.blipEffects = attrs.blipEffects;
  if (attrs.useLocalDpi !== null && attrs.useLocalDpi !== undefined)
    imageOpts.useLocalDpi = attrs.useLocalDpi;
  if (attrs.fill) imageOpts.fill = attrs.fill;
  if (attrs.effects) imageOpts.effects = attrs.effects;
  if (attrs.tile) imageOpts.tile = attrs.tile;
  if (attrs.runPropertiesRawXml) imageOpts.runPropertiesRawXml = attrs.runPropertiesRawXml;

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
  // 0.9.7+ fidelity fields (reverse of renderDocx)
  if (opts.nonVisualProperties) attrs.nonVisualProperties = opts.nonVisualProperties;
  const effectExtent = (opts.transformation as Record<string, unknown> | undefined)?.effectExtent;
  if (effectExtent) attrs.effectExtent = effectExtent;
  if (opts.graphicFrameLocks) attrs.graphicFrameLocks = opts.graphicFrameLocks;
  if (opts.blipEffects) attrs.blipEffects = opts.blipEffects;
  if (opts.useLocalDpi !== undefined) attrs.useLocalDpi = opts.useLocalDpi;
  if (opts.fill) attrs.fill = opts.fill;
  if (opts.effects) attrs.effects = opts.effects;
  if (opts.tile) attrs.tile = opts.tile;
  if (opts.runPropertiesRawXml) attrs.runPropertiesRawXml = opts.runPropertiesRawXml;

  return attrs;
}

// ── Node-level renderHTML helpers ──

/** Dimensions + source needed to size the background image to the extent box. */
export interface CropRenderContext {
  width?: number;
  height?: number;
  src?: string;
}

/**
 * Render crop as background-image styles for byte-accurate four-sided srcRect.
 *
 * object-fit:cover scales uniformly, so it only matches single-axis crops; for
 * two-sided crops it shrinks the whole image and shows too much. background-size
 * scales each axis independently (W/visibleW × H/visibleH), so the visible
 * srcRect region always maps exactly onto the extent box. background-position
 * then shifts the original so the cropped-left/top region falls outside the box.
 *
 * Requires a div[role=img] host (see Image.renderHTML) — an <img> cannot
 * background-crop itself, since its src is always painted in the foreground.
 */
export function renderCropAttrs(
  crop: Record<string, unknown> | CropRect,
  ctx: CropRenderContext = {},
): { style: string } {
  const c = crop as CropRect;
  const leftPct = (c.left || 0) / 100000;
  const topPct = (c.top || 0) / 100000;
  const rightPct = (c.right || 0) / 100000;
  const bottomPct = (c.bottom || 0) / 100000;

  const visibleW = 1 - leftPct - rightPct;
  const visibleH = 1 - topPct - bottomPct;

  // Original displayed size per axis (independent of the other axis — the key
  // difference from object-fit:cover). Falls back to the box size when an axis
  // is uncropped (visible = 1 → same size).
  const w = ctx.width ?? 0;
  const h = ctx.height ?? 0;
  const bgW = visibleW > 0 ? w / visibleW : w;
  const bgH = visibleH > 0 ? h / visibleH : h;

  // background-position % = (box − image) × pct/100; solving for the image's
  // leftPct·bgW point at the box's left edge gives pct = left/(left+right).
  const posX = leftPct + rightPct > 0 ? (leftPct / (leftPct + rightPct)) * 100 : 50;
  const posY = topPct + bottomPct > 0 ? (topPct / (topPct + bottomPct)) * 100 : 50;

  const src = ctx.src ?? "";
  return {
    style: [
      `background-image:url(${src})`,
      `background-size:${bgW.toFixed(2)}px ${bgH.toFixed(2)}px`,
      `background-position:${posX.toFixed(2)}% ${posY.toFixed(2)}%`,
      "background-repeat:no-repeat",
    ].join(";"),
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

  return styles;
}

/**
 * Stamp the nested office-open attrs onto an HTML attribute map as JSON
 * data-* pairs. Shared by the cropped-div and plain-img render branches so the
 * fidelity fields (floating/outline/nonVisualProperties/…) round-trip through
 * HTML identically. `crop` is handled separately (crop branch only).
 */
const RAW_ATTR_DATA: Array<[string, string]> = [
  ["floating", "data-floating"],
  ["outline", "data-outline"],
  ["nonVisualProperties", "data-non-visual"],
  ["effectExtent", "data-effect-extent"],
  ["graphicFrameLocks", "data-graphic-frame-locks"],
  ["blipEffects", "data-blip-effects"],
  ["useLocalDpi", "data-use-local-dpi"],
  ["fill", "data-fill"],
  ["effects", "data-effects"],
  ["tile", "data-tile"],
  ["runPropertiesRawXml", "data-run-properties-raw-xml"],
];

function attachRawAttrs(target: Record<string, unknown>, attrs: Record<string, unknown>): void {
  for (const [attr, data] of RAW_ATTR_DATA) {
    const value = attrs[attr];
    if (value !== null && value !== undefined) target[data] = JSON.stringify(value);
  }
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

      // 0.9.7+ round-trip fidelity fields. office-open parses + stringifies
      // each; we carry them verbatim as JSON in data-* attrs.
      nonVisualProperties: attrDataJson("data-non-visual"), // pic:cNvPr (id/name/descr)
      effectExtent: attrDataJson("data-effect-extent"), // wp:effectExtent (EMUs)
      graphicFrameLocks: attrDataJson("data-graphic-frame-locks"),
      blipEffects: attrDataJson("data-blip-effects"),
      useLocalDpi: attrDataJson("data-use-local-dpi"), // a14:useLocalDpi
      fill: attrDataJson("data-fill"),
      effects: attrDataJson("data-effects"),
      tile: attrDataJson("data-tile"),
      runPropertiesRawXml: attrDataJson("data-run-properties-raw-xml"),
    };
  },

  renderHTML({
    node,
    HTMLAttributes,
  }: {
    node: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) {
    const attrs = node.attrs;

    // Cropped images render as div[role=img] + background-image so background
    // -size can scale each axis independently (object-fit:cover is uniform and
    // only exact for single-axis crops). Non-cropped images stay plain <img>.
    if (attrs.crop) {
      const width = attrs.width as number;
      const height = attrs.height as number;
      const styles = renderImageStyles(attrs);
      const crop = renderCropAttrs(attrs.crop as CropRect, {
        width,
        height,
        src: attrs.src as string | undefined,
      });
      styles.push(crop.style, `width:${width}px`, `height:${height}px`);
      const divAttrs: Record<string, unknown> = {
        "data-image": "crop",
        role: "img",
        style: styles.join(";"),
      };
      if (attrs.alt) divAttrs["aria-label"] = attrs.alt;
      if (attrs.title) divAttrs["title"] = attrs.title;
      attachRawAttrs(divAttrs, attrs);
      divAttrs["data-crop"] = JSON.stringify(attrs.crop);
      return ["div", divAttrs] as const;
    }

    const htmlAttrs: Record<string, unknown> = { ...HTMLAttributes };
    const styles = renderImageStyles(attrs);
    if (styles.length > 0) htmlAttrs.style = styles.join(";");
    attachRawAttrs(htmlAttrs, attrs);
    return ["img", htmlAttrs] as const;
  },

  parseHTML() {
    return [
      {
        tag: "div[data-image=crop]",
        getAttrs: (el) => parseCropDiv(el as HTMLElement),
      },
      { tag: "img[src]" },
    ];
  },

  renderDocx,
  parseDocx,
});

/**
 * Reverse-parse a cropped div[role=img] back into image attrs.
 *
 * Covers only what the attribute-level parseHTML rules cannot recover from a
 * div: src (background-image), width/height (inline style extent box), and
 * alt/title (aria-label/title). rotation/crop/floating/outline are left to
 * their attribute parseHTML rules, which read the style/data-* the div carries.
 */
function parseCropDiv(el: HTMLElement): Record<string, unknown> {
  const style = el.getAttribute("style") || "";
  const attrs: Record<string, unknown> = {};

  // src lives in background-image (the div has no <img> src attribute).
  // Tolerate whitespace around `:` — serializers may emit `background-image: url(..)`.
  const bgMatch = style.match(/background-image:\s*url\(\s*([^)]+?)\s*\)/);
  if (bgMatch) {
    attrs.src = bgMatch[1].replace(/^['"]|['"]$/g, "");
  }

  // width/height live in the inline style (extent box), not HTML attributes
  const wMatch = style.match(/(?:^|;)\s*width:\s*([\d.]+)px/);
  const hMatch = style.match(/(?:^|;)\s*height:\s*([\d.]+)px/);
  if (wMatch) attrs.width = parseFloat(wMatch[1]);
  if (hMatch) attrs.height = parseFloat(hMatch[1]);

  // alt/title carry over via aria-label/title (the div has no alt attribute)
  const ariaLabel = el.getAttribute("aria-label");
  if (ariaLabel) attrs.alt = ariaLabel;
  const title = el.getAttribute("title");
  if (title) attrs.title = title;

  return attrs;
}
