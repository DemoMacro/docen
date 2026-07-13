/**
 * @docen/core/geometry — pure geometry math for OOXML drawings.
 *
 * All functions are stateless and carry no DOM/LeaferJS dependency, so they run
 * identically in the browser and in Node (SSR / headless export). EMU/px
 * conversion delegates to `@office-open/core/util` so there is one source of
 * truth for the EMU-per-pixel constant.
 *
 * @module
 */

import { convertEmuToPixels, convertToEmu } from "@office-open/core/util";

/** Per-axis visible fraction produced by {@link cropFractions}. */
export interface CropFractions {
  /** Visible width fraction (0–1); 1 when horizontally uncropped. */
  visibleW: number;
  /** Visible height fraction (0–1); 1 when vertically uncropped. */
  visibleH: number;
  /** Left offset fraction (0–1) of the cropped-out region. */
  left: number;
  /** Top offset fraction (0–1) of the cropped-out region. */
  top: number;
}

/** Convert OOXML permyriad (0–100000) to a 0–1 fraction. */
export const permyriadToFraction = (value: number | undefined): number => (value ?? 0) / 100000;

/** True when any side of an OOXML srcRect is non-zero. */
export const hasCrop = (
  crop: { left?: number; top?: number; right?: number; bottom?: number } | null | undefined,
): boolean => Boolean(crop && (crop.left || crop.top || crop.right || crop.bottom));

/**
 * Resolve an OOXML srcRect (permyriad per side) into per-axis visible/offset
 * fractions.
 *
 * Mirrors the math in `@docen/docx` `renderCropAttrs` so the canvas crop box
 * matches the HTML crop box byte-for-byte. Returns `{ visibleW:1, visibleH:1,
 * left:0, top:0 }` when nothing is cropped.
 */
export const cropFractions = (
  crop: { left?: number; top?: number; right?: number; bottom?: number } | null | undefined,
): CropFractions => {
  if (!crop) return { visibleW: 1, visibleH: 1, left: 0, top: 0 };
  const left = permyriadToFraction(crop.left);
  const top = permyriadToFraction(crop.top);
  const right = permyriadToFraction(crop.right);
  const bottom = permyriadToFraction(crop.bottom);
  const visibleW = Math.max(0, 1 - left - right);
  const visibleH = Math.max(0, 1 - top - bottom);
  return { visibleW, visibleH, left, top };
};

/** Convert an EMU value to CSS pixels (rounded for stable layout). */
export const emuToPx = (emu: number | undefined): number =>
  Math.round(convertEmuToPixels(emu ?? 0));

/** Convert CSS pixels to EMU (for writing geometry back to OOXML). */
export const pxToEmu = (px: number): number => convertToEmu(px);

/** Clamp a dimension to a positive, finite number, falling back to `fallback`. */
export const clampDimension = (value: number | null | undefined, fallback: number): number =>
  typeof value === "number" && value > 0 && Number.isFinite(value) ? value : fallback;
