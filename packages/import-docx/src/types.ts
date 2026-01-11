/**
 * DOCX image information for custom converter
 */
export interface DocxImageInfo {
  /** Image ID (relationship ID in DOCX) */
  id: string;
  /** Content type (e.g., "image/png", "image/jpeg") */
  contentType: string;
  /** Raw image data */
  data: Uint8Array;
}

/**
 * Result of image conversion
 */
export interface DocxImageResult {
  /** Image src attribute value (URL or data URL) */
  src: string;
  /** Optional alt text */
  alt?: string;
}

/**
 * Custom image converter function type
 */
export type DocxImageConverter = (image: DocxImageInfo) => Promise<DocxImageResult>;
