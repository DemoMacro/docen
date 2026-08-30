import { encodeBase64 } from "@office-open/core";
import { convertEmuToPixels } from "@office-open/core/util";
import type { ParagraphChild, PictureOptions } from "@office-open/docx";
import { Node } from "@tiptap/core";

import type { JSONContent } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/** The picture ParagraphChild branch office-open parses a drawing run into. */
type PictureBranch = Extract<ParagraphChild, { picture: PictureOptions }>;

/**
 * Custom Image extension with renderDocx/parseDocx (rendering is canvas-only;
 * HTML is parse-input only).
 *
 * Attrs:
 *  - src/alt/title/width/height: Tiptap structural names (kept verbatim so base
 *    image commands work).
 *  - rotation: carried through DOCX via transformation.rotation
 *    (MediaTransformation.rotation).
 *  - flipH/flipV: DOCX round-trip via transformation.flipHorizontal/flipVertical
 *    (a:xfrm@flipH/flipV). Three-state (null/true/false) — an explicit false
 *    emits flipH="0" byte-faithfully.
 *  - floating/outline: nested office-open objects (Floating / OutlineOptions).
 *  - crop: nested office-open SourceRectangleOptions (srcRect).
 *  - display: editor-only display hint, no OOXML equivalent.
 *
 * DOCX round-trip is near-identity: renderDocx packs attrs into CorePictureOptions;
 * parseDocx unpacks them back. src is a data URL ↔ { type, data } base64.
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

/** Decoded src cache, keyed by the node's attrs object itself: Node.toJSON
 *  carries attrs by reference, so untouched images keep a stable identity
 *  across transactions and the megabyte atob runs once per image, not once
 *  per keystroke. A src-content Map thrashes here instead — the corpus can
 *  hold more images than any bounded entry count, and content-hashing
 *  megabyte keys each pass costs as much as the decode. When the editor
 *  rewrites the attrs (image replaced), the old object dies and the entry
 *  is collected with it. Callers must not mutate the returned arrays — the
 *  cache hands out shared instances. */
const decodedByAttrs = new WeakMap<object, Uint8Array>();

function decodedBytesOf(attrs: object, src: string): Uint8Array | undefined {
  const hit = decodedByAttrs.get(attrs);
  if (hit) return hit;
  const comma = src.indexOf(",");
  if (comma < 0) return undefined;
  try {
    const bin = atob(src.slice(comma + 1));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    decodedByAttrs.set(attrs, bytes);
    return bytes;
  } catch {
    return undefined;
  }
}

/**
 * Tiptap JSON image node → CorePictureOptions-shaped object.
 *
 * Returns `{ picture: PictureOptions }` (structural wrapper) or null when no
 * embedded image data is available (external URLs need pre-fetching).
 * rotation is carried via transformation.rotation (not dropped).
 */
export function renderDocx(node: JSONContent): Record<string, unknown> | null {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const imageOpts: Record<string, unknown> = {};

  // src (data URL) → { type, data }, decoded through the shared cache so the
  // projection downstream sees a stable bytes identity across transactions.
  const src = attrs.src as string | undefined;
  if (src?.startsWith("data:image/")) {
    const match = src.match(/^data:image\/([\w.+-]+);base64,/);
    if (match) {
      imageOpts.type = match[1] === "jpeg" ? "jpg" : match[1];
      imageOpts.data = decodedBytesOf(attrs, src);
    }
  }

  // Cannot generate an image run without embedded data (external URLs need pre-fetching)
  if (!imageOpts.data) return null;

  // transformation: width/height are required by OOXML MediaTransformation —
  // default when absent (editor/prepare step normally supplies real dimensions).
  // rotation is an optional editor attr carried via transformation.rotation.
  // Guard against NaN/non-finite reaching OOXML as `${width}px` (an invalid
  // UniversalMeasure that corrupts the document). Falls back to the default.
  const width = Number.isFinite(attrs.width as number) ? (attrs.width as number) : 400;
  const height = Number.isFinite(attrs.height as number) ? (attrs.height as number) : 300;
  // office-open 0.10.4+ treats a numeric transformation size as EMU (was px);
  // emit UniversalMeasure so the px value is interpreted correctly on generate.
  const transformation: Record<string, unknown> = { width: `${width}px`, height: `${height}px` };
  const rotation = attrs.rotation as number | undefined;
  if (rotation != null) transformation.rotation = rotation;
  // flip: office-open MediaTransformation.flip is {horizontal, vertical} (mapped
  // to a:xfrm flipH/flipV by createTransformation). Three-state per axis: null =
  // omit, true/false both emit (office-open keeps an explicit flipH="0" — byte-
  // faithful round-trip needs the false case). Build the object only when at
  // least one axis is set.
  const flipHSet = attrs.flipH !== null && attrs.flipH !== undefined;
  const flipVSet = attrs.flipV !== null && attrs.flipV !== undefined;
  if (flipHSet || flipVSet) {
    transformation.flip = {
      ...(flipHSet ? { horizontal: attrs.flipH as boolean } : {}),
      ...(flipVSet ? { vertical: attrs.flipV as boolean } : {}),
    };
  }
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
  if (attrs.runProperties) imageOpts.runProperties = attrs.runProperties;

  return { picture: imageOpts };
}

/**
 * PictureOptions-shaped object → Tiptap attrs.
 *
 * Near-identity unpack: transformation → width/height/rotation, altText → alt/title,
 * floating/srcRect(→crop)/outline passed through verbatim. src is reconstructed by
 * DocxManager from the image data bytes (kept out of parseDocx).
 */
export function parseDocx(picture: PictureOptions): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};

  // transformation → width/height/rotation (structural Tiptap attrs)
  const { transformation } = picture;
  if (transformation) {
    // office-open 0.10.4+ parses wp:extent as EMU verbatim (was px); convert to px.
    if (typeof transformation.width === "number")
      attrs.width = convertEmuToPixels(transformation.width);
    if (typeof transformation.height === "number")
      attrs.height = convertEmuToPixels(transformation.height);
    if (typeof transformation.rotation === "number") attrs.rotation = transformation.rotation;
    // office-open MediaTransformation.flip is {horizontal, vertical}; carry both
    // true and false through per axis (false is meaningful: flipH="0").
    const { flip } = transformation;
    if (flip) {
      if (flip.horizontal !== undefined) attrs.flipH = flip.horizontal;
      if (flip.vertical !== undefined) attrs.flipV = flip.vertical;
    }
  }

  // altText → alt/title
  const { altText } = picture;
  if (altText) {
    if (altText.name) attrs.alt = altText.name;
    if (altText.description) attrs.title = altText.description;
  }

  // Near-identity pass-through for nested office-open objects
  if (picture.floating) attrs.floating = picture.floating;
  if (picture.sourceRectangle) attrs.crop = picture.sourceRectangle;
  if (picture.outline) attrs.outline = picture.outline;
  // 0.9.7+ fidelity fields (reverse of renderDocx)
  if (picture.nonVisualProperties) attrs.nonVisualProperties = picture.nonVisualProperties;
  if (transformation?.effectExtent) attrs.effectExtent = transformation.effectExtent;
  if (picture.graphicFrameLocks) attrs.graphicFrameLocks = picture.graphicFrameLocks;
  if (picture.blipEffects) attrs.blipEffects = picture.blipEffects;
  if (picture.useLocalDpi !== undefined) attrs.useLocalDpi = picture.useLocalDpi;
  if (picture.fill) attrs.fill = picture.fill;
  if (picture.effects) attrs.effects = picture.effects;
  if (picture.tile) attrs.tile = picture.tile;
  if (picture.runProperties) attrs.runProperties = picture.runProperties;

  return attrs;
}

/** ParagraphChild `{ picture: PictureOptions }` → image node. Mirrors the old
 *  DocxManager.resolveImage: reflective attrs parse, then rebuild the data URL
 *  from the embedded bytes (encodeBase64 handles platform dispatch + stack
 *  guard). */
function resolveImage(picture: PictureOptions, ctx: ResolveContext): JSONContent {
  const attrs = ctx.parseNodeAttrs("image", picture);
  const { data, type } = picture;
  // office-open parse always yields bytes; guard the other DataType members
  // (string/ArrayBuffer) so the data URL is built from Uint8Array only.
  const bytes =
    data instanceof Uint8Array ? data : data instanceof ArrayBuffer ? new Uint8Array(data) : null;
  if (bytes && type) {
    attrs.src = `data:image/${type};base64,${encodeBase64(bytes)}`;
  }
  return { type: "image", attrs };
}

// DOCX image run → office-open ParagraphChild `{ picture: PictureOptions }`.
export const parseDocxInline: ParseInlineRule<PictureBranch> = {
  match: (child): child is PictureBranch => "picture" in child,
  convert: (child, ctx) => resolveImage(child.picture, ctx),
};

// ── Extension ──

/** Fully custom image node (no upstream extension): an inline atom carrying
 *  src/alt/title plus the office-open mirror below. */
export const Image = Node.create({
  name: "image",
  inline: true,
  group: "inline",
  draggable: true,

  addAttributes() {
    return {
      src: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute("src"),
      },
      alt: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute("alt"),
      },
      title: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute("title"),
      },

      // Editor-only display hint (no OOXML equivalent; not round-tripped)
      display: {
        default: null,
        rendered: false,
        parseHTML: () => "inline-block",
      },

      // Editor-only loading state for an unsized http image (image-cap stamps
      // loading/error/timeout). A transient runtime state — never present in
      // imported HTML/JSON — so no parseHTML rule and not round-tripped.
      loadState: {
        default: null,
        rendered: false,
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

      // DOCX transformation.flipHorizontal/flipVertical (a:xfrm@flipH/flipV).
      // Three-state: null = omit (default), true/false = emit flipH="1"/"0".
      // office-open keeps the explicit false, so byte-faithful round-trip needs
      // null vs false distinct (rotation has no such case — it's a number).
      flipH: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const style = element.getAttribute("style") || "";
          // Only match scaleX(-1) inside a transform list (not a translate or
          // matrix that happens to contain -1); anchored on the scale token.
          return /(^|[\s(])scaleX\(-1\)/.test(style) || null;
        },
      },
      flipV: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => {
          const style = element.getAttribute("style") || "";
          return /(^|[\s(])scaleY\(-1\)/.test(style) || null;
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

      // Nested office-open Floating (JSON in data-floating)
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
      runProperties: attrDataJson("data-run-properties"),
    };
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
