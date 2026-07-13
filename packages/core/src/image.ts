/**
 * @docen/core/image — OOXML image data ↔ LeaferJS Image element.
 *
 * The canonical image shape inside an Office document is a DrawingML `<pic:pic>`
 * (a `<a:blip>` wrapped in a transform, optional srcRect crop, outline, and
 * floating anchor). `@docen/docx` carries the round-trip schema; this module
 * turns the resulting {@link RenderImageInput} into LeaferJS `IImageInputData`
 * (the native constructor options for an `Image` element), ready for the
 * editor's `<docen-image>` component to mount, and turns a user-resized
 * element back into the same data shape for DOCX persistence.
 *
 * The functions are DOM-free apart from the LeaferJS types they reference, so a
 * Node-side caller (`@leafer-ui/node`) can render the exact same element for
 * headless export / thumbnails.
 *
 * @module
 */

import type { OutlineOptions, SourceRectangleOptions } from "@office-open/core/drawingml";
import type { IImageInputData, IUI } from "leafer-ui";

import { clampDimension, cropFractions } from "./geometry";
import { renderOutline } from "./style";

/** Minimal subset of `ImageAttrs` (from `@docen/docx`) that the renderer needs.
 *  Declared locally so @docen/core has no hard dependency on the docx engine —
 *  pptx/xlsx image data can conform to the same shape. The `crop`/`outline`
 *  fields reuse `@office-open/core`'s native types (`SourceRectangleOptions` /
 *  `OutlineOptions`) instead of redefining them. */
export interface RenderImageInput {
  /** Image source as a URL or `data:` URL. */
  src: string;
  /** X position in CSS pixels (defaults to 0). */
  x?: number | null;
  /** Y position in CSS pixels (defaults to 0). */
  y?: number | null;
  /** Display width in CSS pixels (OOXML extent already converted). */
  width: number | null | undefined;
  /** Display height in CSS pixels (OOXML extent already converted). */
  height: number | null | undefined;
  /** Clockwise rotation in degrees (OOXML `transformation.rotation`). */
  rotation?: number | null;
  /** OOXML srcRect crop (`a:srcRect`), permyriad per side. */
  crop?: SourceRectangleOptions | null;
  /** OOXML outline (`a:ln`). */
  outline?: OutlineOptions | null;
}

/** Geometric fields the editor reads back after a user move/resize/rotate. */
export interface ParsedImageOutput {
  /** X position in CSS pixels (LeaferJS local element.x). */
  x: number;
  /** Y position in CSS pixels (LeaferJS local element.y). */
  y: number;
  /** New display width in CSS pixels. */
  width: number;
  /** New display height in CSS pixels. */
  height: number;
  /** New clockwise rotation in degrees, normalized to [0, 360). */
  rotation: number;
}

/** Default display size when width/height are missing or invalid (matches the
 *  `@docen/docx` image renderHTML placeholder ratio 4:3). */
const DEFAULT_WIDTH = 400;
const DEFAULT_HEIGHT = 300;

/** Normalize a rotation in degrees to [0, 360). */
const normalizeRotation = (deg: number | null | undefined): number => {
  if (!deg || !Number.isFinite(deg)) return 0;
  return ((deg % 360) + 360) % 360;
};

/**
 * Build the LeaferJS `Image` element options (`IImageInputData`) for an OOXML
 * image.
 *
 * Width / height default to 400×300 when absent (the same placeholder ratio the
 * engine's `renderHTML` uses). Rotation is normalized to [0, 360). The OOXML
 * outline is mapped to LeaferJS `stroke` / `strokeWidth` / `dashPattern` via
 * {@link renderOutline}. srcRect crop is NOT included here (LeaferJS's Image
 * ignores an unknown `crop` key) — the editor component reads `input.crop`
 * and wraps the image in a Box{overflow:'hide'} sized via {@link renderCropBox}.
 */
export const renderImage = (input: RenderImageInput): IImageInputData => {
  const width = clampDimension(input.width, DEFAULT_WIDTH);
  const height = clampDimension(input.height, DEFAULT_HEIGHT);
  const rotation = normalizeRotation(input.rotation);
  const outline = renderOutline(input.outline);
  const options: IImageInputData = {
    url: input.src,
    x: input.x ?? 0,
    y: input.y ?? 0,
    width,
    height,
    rotation,
    // No origin/around: resize uses the top-left as anchor (Office behavior —
    // drag the bottom-right handle, the top-left stays put). Rotation/flip use
    // rotateOf('center') / flip() which have their own center-based pivot, so
    // they don't need a persistent origin on the element.
    // editable: true is required by @leafer-in/editor for an element to be
    // selectable / resizable / rotatable (without it the editor ignores it).
    editable: true,
  };
  if (outline) {
    options.stroke = outline.stroke;
    options.strokeWidth = outline.strokeWidth;
    options.dashPattern = outline.dashPattern;
  }
  return options;
};

/** Read the user-moved/resized geometry of a LeaferJS element back into the
 *  OOXML data shape. Use this in the editor's transform-change handler.
 *
 *  Reads LOCAL geometry (element.x/y/width/height) rather than `boxBounds`:
 *  boxBounds is the world-space axis-aligned bounding box, which swaps/enlarges
 *  under rotation — reading it after rotate() would bleed the rotated AABB back
 *  into width/height and corrupt the exported extent. Local width/height are
 *  unchanged by rotation (only `rotation` changes), so a pure rotate() round-
 *  trips without dimension drift (edit == render == export). */
export const parseImage = (element: IUI): ParsedImageOutput => ({
  x: Math.round(element.x ?? 0),
  y: Math.round(element.y ?? 0),
  width: Math.round(element.width ?? 0),
  height: Math.round(element.height ?? 0),
  rotation: normalizeRotation(element.rotation ?? 0),
});

/** Compute the LeaferJS inner-image size + offset that realize a srcRect crop
 *  inside an outer box of the given width/height. Used by the editor component
 *  to lay out a crop frame (mirrors `@docen/docx` `renderCropAttrs` math). */
export const renderCropBox = (
  crop: SourceRectangleOptions,
  width: number,
  height: number,
): { innerWidth: number; innerHeight: number; offsetX: number; offsetY: number } => {
  const { visibleW, visibleH, left, top } = cropFractions(crop);
  const innerWidth = visibleW > 0 ? width / visibleW : width;
  const innerHeight = visibleH > 0 ? height / visibleH : height;
  return {
    innerWidth,
    innerHeight,
    offsetX: -(left * innerWidth),
    offsetY: -(top * innerHeight),
  };
};
