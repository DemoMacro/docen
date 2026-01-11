import type { DocxImageConverter } from "./types";

/**
 * Options for importing DOCX files to TipTap content
 */
export interface DocxImportOptions {
  /** Custom image converter (default: embed as base64 data URL) */
  convertImage?: DocxImageConverter;

  /** Whether to ignore empty paragraphs (default: false) */
  ignoreEmptyParagraphs?: boolean;
}
