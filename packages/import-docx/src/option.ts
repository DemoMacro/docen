import type { DocxImageConverter } from "./types";

/**
 * Options for importing DOCX files to TipTap content
 */
export interface DocxImportOptions {
  /** Custom image converter (default: embed as base64 data URL) */
  convertImage?: DocxImageConverter;

  /** Whether to ignore empty paragraphs (default: false) */
  ignoreEmptyParagraphs?: boolean;

  /**
   * Dynamic import function for @napi-rs/canvas
   * Required for image cropping in Node.js environment, ignored in browser
   *
   * @example
   * import { parseDOCX } from '@docen/import-docx';
   * const content = await parseDOCX(buffer, {
   *   canvasImport: () => import('@napi-rs/canvas')
   * });
   */
  canvasImport?: () => Promise<typeof import("@napi-rs/canvas")>;

  /**
   * Enable or disable image cropping during import
   * When true (default), images with crop information in DOCX will be cropped
   * When false, crop information is ignored and full image is used
   *
   * @default true
   */
  enableImageCrop?: boolean;
}
