/**
 * @docen/core/export — canvas → image (base64 / bytes) via LeaferJS export.
 *
 * Thin wrapper over the LeaferJS `@leafer-in/export` plugin so callers stay
 * decoupled from the LeaferJS element API. Any LeaferJS element (`App`,
 * `Leafer`, `Group`, `Image`, `UI`) that has been augmented by the export plugin
 * exposes `export()` / `syncExport()`; these helpers forward to those methods
 * and normalize the `IExportResult.data` into a base64 data URL.
 *
 * @module
 */

import type { IExportOptions, IExportResult, IUI } from "leafer-ui";

/** Format accepted by {@link exportImage} / {@link exportCanvas}. */
export type ExportFormat = "png" | "jpg" | "jpeg";

/** Options forwarded to LeaferJS `element.export()`. */
export interface ExportRequestOptions {
  /** Pixel ratio for high-DPI export (e.g. 2 for retina). Default 1. */
  pixelRatio?: number;
  /** JPEG quality 0–1 (ignored for png). Default 0.92. */
  quality?: number;
}

/** Build the LeaferJS `IExportOptions` from our simpler {@link
 *  ExportRequestOptions}. */
const toExportOptions = (options: ExportRequestOptions): IExportOptions =>
  ({
    pixelRatio: options.pixelRatio ?? 1,
    quality: options.quality ?? 0.92,
  }) as IExportOptions;

/** Coerce a LeaferJS export result's `data` field to a base64 data URL string.
 *  Returns "" when the plugin returned a non-string (canvas / blob) — callers
 *  should then use the element's `export(filename)` form for file output. */
const toDataUrl = (data: IExportResult["data"]): string => (typeof data === "string" ? data : "");

/**
 * Export a LeaferJS element to a base64 data URL.
 *
 * The element must come from a LeaferJS app that has loaded `@leafer-in/export`
 * (the `@docen/editor` image/shape/chart editor components set this up). Returns
 * a `data:image/<format>;base64,…` string ready to drop into an `<img src>` or
 * an OOXML `<a:blip>` payload.
 */
export const exportImage = async (
  element: IUI,
  format: ExportFormat = "png",
  options: ExportRequestOptions = {},
): Promise<string> => {
  const fmt = format === "jpeg" ? "jpg" : format;
  const result = await element.export(fmt, toExportOptions(options));
  return toDataUrl(result.data);
};

/**
 * Export a LeaferJS element synchronously.
 *
 * Only works when all async-loaded images in the element have finished decoding
 * (the caller's responsibility). Prefer {@link exportImage} unless you are on a
 * hot path that has already awaited image loads.
 */
export const exportCanvas = (
  element: IUI,
  format: ExportFormat = "png",
  options: ExportRequestOptions = {},
): string => {
  const fmt = format === "jpeg" ? "jpg" : format;
  const result = element.syncExport(fmt, toExportOptions(options));
  return toDataUrl(result.data);
};
