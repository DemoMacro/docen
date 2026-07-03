import { convertEmuToPixels, encodeBase64 } from "@office-open/core";
import { Image as BaseImage } from "@tiptap/extension-image";

import type { JSONContent } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";
import { floatAnchorScope, floatingToStyles } from "./utils";

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
  // Guard against NaN/non-finite reaching OOXML as `${width}px` (an invalid
  // UniversalMeasure that corrupts the document). Falls back to the default.
  const width = Number.isFinite(attrs.width as number) ? (attrs.width as number) : 600;
  const height = Number.isFinite(attrs.height as number) ? (attrs.height as number) : 400;
  // office-open 0.10.4+ treats a numeric transformation size as EMU (was px);
  // emit UniversalMeasure so the px value is interpreted correctly on generate.
  const transformation: Record<string, unknown> = { width: `${width}px`, height: `${height}px` };
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
  if (attrs.crop) imageOpts.sourceRectangle = attrs.crop;
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
    // office-open 0.10.4+ parses wp:extent as EMU verbatim (was px); convert to px.
    if (typeof transformation.width === "number")
      attrs.width = convertEmuToPixels(transformation.width);
    if (typeof transformation.height === "number")
      attrs.height = convertEmuToPixels(transformation.height);
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
  if (opts.sourceRectangle) attrs.crop = opts.sourceRectangle;
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

/** ParagraphChild `{ image: ImageOptions }` → image node. Mirrors the old
 *  DocxManager.resolveImage: reflective attrs parse, then rebuild the data URL
 *  from the embedded bytes (encodeBase64 handles platform dispatch + stack
 *  guard). */
function resolveImage(imageOpts: Record<string, unknown>, ctx: ResolveContext): JSONContent {
  const attrs = ctx.parseNodeAttrs("image", imageOpts);
  const data = imageOpts.data as Uint8Array | undefined;
  const type = imageOpts.type as string | undefined;
  if (data && type) {
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    attrs.src = `data:image/${type};base64,${encodeBase64(bytes)}`;
  }
  return { type: "image", attrs };
}

// DOCX image run → office-open ParagraphChild `{ image: ImageOptions }`.
export const parseDocxInline: ParseInlineRule = {
  match: (child) => "image" in child,
  convert: (child, ctx) => resolveImage((child as { image: Record<string, unknown> }).image, ctx),
};

// ── Node-level renderHTML helpers ──

/** Extent-box dimensions needed to size the inner <img> for a cropped image. */
export interface CropRenderContext {
  width?: number;
  height?: number;
}

/**
 * Render crop as the inner <img> style for byte-accurate four-sided srcRect.
 *
 * object-fit:cover scales uniformly, so it only matches single-axis crops.
 * Instead the inner <img> is sized to the un-cropped display size per axis
 * (imgW = W/visibleW, imgH = H/visibleH) and translated so the visible srcRect
 * region maps exactly onto the extent box; the outer box clips (overflow:hidden)
 * the cropped-out left/top region. Mathematically equivalent to a
 * background-size/background-position crop, but keeps a real <img> (alt,
 * accessibility, semantics, drag-to-save).
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

  // Inner <img> display size per axis (independent of the other axis — the key
  // difference from object-fit:cover). Falls back to the box size when an axis
  // is uncropped (visible = 1 → same size).
  const w = ctx.width ?? 0;
  const h = ctx.height ?? 0;
  const imgW = visibleW > 0 ? w / visibleW : w;
  const imgH = visibleH > 0 ? h / visibleH : h;

  // Shift the image so the visible region's top-left lands at the box's
  // top-left: translate by the cropped-left/top amount (pct × display size).
  const offX = -(leftPct * imgW);
  const offY = -(topPct * imgH);

  return {
    style: [
      "display:block",
      `width:${imgW.toFixed(2)}px`,
      `height:${imgH.toFixed(2)}px`,
      `transform:translate(${offX.toFixed(2)}px,${offY.toFixed(2)}px)`,
      "transform-origin:0 0",
    ].join(";"),
  };
}

function renderImageStyles(attrs: Record<string, unknown>): string[] {
  const styles: string[] = [];

  if (attrs.display) {
    styles.push(`display:${attrs.display as string}`);
  } else if (!attrs.floating) {
    // Cropped images render as a <span> box and vector images as a <div> — both
    // are block-level by default, so without inline-block a box claims its own
    // line and breaks side-by-side image rows (a centered row of two images
    // collapses to one-per-line). floating images are placed via
    // position:absolute/float, where display no longer governs layout.
    styles.push("display:inline-block");
  }

  if (attrs.rotation) {
    styles.push(`transform:rotate(${attrs.rotation as number}deg)`);
  }

  if (attrs.floating) {
    // Floating anchor (wp:anchor) → CSS. Shared with WpgGroup via utils so
    // anchored images and drawing groups render identically (wrapNone →
    // position:absolute "in front/behind text"; square/tight → float).
    styles.push(
      ...floatingToStyles(
        attrs.floating,
        attrs.src as string | undefined,
        attrs.width as number | undefined,
      ),
    );
  }

  return styles;
}

/** Office vector (GDI) formats the browser cannot decode — EMF/WMF render as
 *  a labeled placeholder (Image.renderHTML) instead of an <img> that decodes
 *  to naturalWidth 0 (an empty box). */
function isVectorImage(src: unknown): boolean {
  return typeof src === "string" && /^data:image\/(?:x-)?(?:emf|wmf)/i.test(src);
}

/** Remote http(s) images are lazy-loaded + async-decoded in renderHTML; data
 *  URLs (the imported-DOCX majority) have no network to defer, so they stay
 *  eager — and their load event, awaited by the editor's cap path, is not
 *  delayed by `loading="lazy"`. */
function isRemoteImage(src: unknown): boolean {
  return typeof src === "string" && /^https?:/.test(src);
}

/** Render attrs applied only to remote images (see isRemoteImage). */
const REMOTE_IMG_ATTRS: Record<string, unknown> = { loading: "lazy", decoding: "async" };

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

      // width/height: capture from the HTML attribute (default) OR inline style
      // (px). External HTML — especially pasted from Word/web — usually sizes
      // images via style="width:..px" rather than a width=".." attribute, so
      // reading the style keeps sizing through HTML→JSON parsing.
      width: {
        parseHTML: (element: HTMLElement) => {
          const attr = element.getAttribute("width");
          if (attr != null) {
            const n = parseFloat(attr);
            if (!Number.isNaN(n)) return n;
          }
          const style = element.getAttribute("style") || "";
          const m = style.match(/(?:^|;)\s*width:\s*([\d.]+)px/);
          return m ? parseFloat(m[1]) : null;
        },
      },
      height: {
        parseHTML: (element: HTMLElement) => {
          const attr = element.getAttribute("height");
          if (attr != null) {
            const n = parseFloat(attr);
            if (!Number.isNaN(n)) return n;
          }
          const style = element.getAttribute("style") || "";
          const m = style.match(/(?:^|;)\s*height:\s*([\d.]+)px/);
          return m ? parseFloat(m[1]) : null;
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
    // A paragraph-anchored wrapNone image resolves its absolute top/left from
    // the anchor <p> (data-float-anchor → editor CSS makes the <p> relative);
    // otherwise it anchors to the page box and floats over the heading/body.
    const floatAnchor =
      attrs.floating && floatAnchorScope(attrs.floating) === "paragraph" ? "paragraph" : null;

    // EMF/WMF are Office GDI vector formats the browser cannot decode — the
    // <img> yields naturalWidth 0 (an empty box). Render a labeled placeholder
    // that keeps the image's size/floating/rotation so pagination and float
    // anchoring stay faithful; the real art is preserved in data-vector-src and
    // the node attrs for DOCX round-trip.
    if (isVectorImage(attrs.src)) {
      const styles = renderImageStyles(attrs);
      if (attrs.width != null) styles.push(`width:${attrs.width as number}px`);
      if (attrs.height != null) styles.push(`height:${attrs.height as number}px`);
      const divAttrs: Record<string, unknown> = {
        "data-image": "vector",
        role: "img",
        "data-vector-src": attrs.src as string,
        style: styles.join(";"),
      };
      if (attrs.alt) divAttrs["aria-label"] = attrs.alt;
      if (attrs.title) divAttrs["title"] = attrs.title;
      attachRawAttrs(divAttrs, attrs);
      if (floatAnchor) divAttrs["data-float-anchor"] = floatAnchor;
      return ["div", divAttrs, "Vector image"] as const;
    }

    // Cropped images render as a span[extent-box] (overflow:hidden + placement)
    // wrapping a real <img>. The <img> is sized per axis and translated so the
    // visible srcRect region maps onto the box (object-fit:cover is uniform and
    // only exact for single-axis crops); the box clips the cropped-out region.
    // Keeping a real <img> (vs a background-image div) preserves alt/accessibility.
    if (attrs.crop) {
      const width = attrs.width as number;
      const height = attrs.height as number;
      const boxStyles = renderImageStyles(attrs);
      boxStyles.push("overflow:hidden", `width:${width}px`, `height:${height}px`);
      const crop = renderCropAttrs(attrs.crop as CropRect, { width, height });
      const boxAttrs: Record<string, unknown> = {
        "data-image": "crop",
        style: boxStyles.join(";"),
      };
      attachRawAttrs(boxAttrs, attrs);
      boxAttrs["data-crop"] = JSON.stringify(attrs.crop);
      if (floatAnchor) boxAttrs["data-float-anchor"] = floatAnchor;
      const imgAttrs: Record<string, unknown> = {
        src: attrs.src as string,
        style: crop.style,
        ...(isRemoteImage(attrs.src) ? REMOTE_IMG_ATTRS : {}),
      };
      if (attrs.alt) imgAttrs.alt = attrs.alt;
      if (attrs.title) imgAttrs.title = attrs.title;
      return ["span", boxAttrs, ["img", imgAttrs]] as const;
    }

    const htmlAttrs: Record<string, unknown> = {
      ...HTMLAttributes,
      ...(isRemoteImage(attrs.src) ? REMOTE_IMG_ATTRS : {}),
    };
    const styles = renderImageStyles(attrs);
    if (styles.length > 0) htmlAttrs.style = styles.join(";");
    attachRawAttrs(htmlAttrs, attrs);
    if (floatAnchor) htmlAttrs["data-float-anchor"] = floatAnchor;
    return ["img", htmlAttrs] as const;
  },

  parseHTML() {
    return [
      {
        tag: "span[data-image=crop]",
        getAttrs: (el) => parseCropDiv(el as HTMLElement),
      },
      {
        tag: "div[data-image=vector]",
        getAttrs: (el) => parseVectorDiv(el as HTMLElement),
      },
      { tag: "img[src]" },
    ];
  },

  renderDocx,
  parseDocx,
  parseDocxInline,
});

/**
 * Reverse-parse a cropped span[extent-box] back into image attrs.
 *
 * src/alt/title live on the inner <img>; width/height on the outer box inline
 * style (the extent box, not the img's un-cropped display size).
 * rotation/crop/floating/outline are left to their attribute parseHTML rules,
 * which read the style/data-* the box carries.
 */
function parseCropDiv(el: HTMLElement): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};

  // src/alt/title live on the inner <img> (the box carries none of them).
  const img = el.querySelector("img");
  if (img) {
    const src = img.getAttribute("src");
    if (src) attrs.src = src;
    const alt = img.getAttribute("alt");
    if (alt) attrs.alt = alt;
    const title = img.getAttribute("title");
    if (title) attrs.title = title;
  }

  // width/height live in the box inline style (extent), not the <img> (which
  // carries the larger un-cropped display size + transform).
  const style = el.getAttribute("style") || "";
  const wMatch = style.match(/(?:^|;)\s*width:\s*([\d.]+)px/);
  const hMatch = style.match(/(?:^|;)\s*height:\s*([\d.]+)px/);
  if (wMatch) attrs.width = parseFloat(wMatch[1]);
  if (hMatch) attrs.height = parseFloat(hMatch[1]);

  return attrs;
}

/** Reverse-parse an EMF/WMF placeholder div back into image attrs: src from
 *  data-vector-src, extent from the inline width/height, alt/title from
 *  aria-label/title. Floating/rotation/etc. round-trip via attribute rules. */
function parseVectorDiv(el: HTMLElement): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  const src = el.getAttribute("data-vector-src");
  if (src) attrs.src = src;
  const style = el.getAttribute("style") || "";
  const wMatch = style.match(/(?:^|;)\s*width:\s*([\d.]+)px/);
  const hMatch = style.match(/(?:^|;)\s*height:\s*([\d.]+)px/);
  if (wMatch) attrs.width = parseFloat(wMatch[1]);
  if (hMatch) attrs.height = parseFloat(hMatch[1]);
  const ariaLabel = el.getAttribute("aria-label");
  if (ariaLabel) attrs.alt = ariaLabel;
  const title = el.getAttribute("title");
  if (title) attrs.title = title;
  return attrs;
}
