import type { DocxImageImportHandler } from "./types";

// Options for importing DOCX files to TipTap content
export interface DocxImportOptions {
  image?: {
    // Custom image handler (default: embed as base64 data URL)
    handler?: DocxImageImportHandler;

    // Dynamic import function for @napi-rs/canvas
    // Required for image cropping in Node.js environment, ignored in browser
    canvasImport?: () => Promise<typeof import("@napi-rs/canvas")>;

    // Enable or disable image cropping during import
    // When true, images with crop information in DOCX will be cropped
    // When false (default), crop information is ignored and full image is used
    enableImageCrop?: boolean;
  };

  // Whether to ignore empty paragraphs (default: false)
  ignoreEmptyParagraphs?: boolean;
}
