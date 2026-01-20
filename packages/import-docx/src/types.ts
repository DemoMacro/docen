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

/**
 * List information extracted from numbering.xml
 */
export interface ListInfo {
  type: "bullet" | "ordered";
  start?: number;
}

/**
 * Map of numbering ID to list information
 */
export type ListTypeMap = Map<string, ListInfo>;

/**
 * Image information with dimensions (for round-trip conversion)
 */
export interface ImageInfo {
  src: string; // data URL (e.g., "data:image/png;base64,...")
  width?: number;
  height?: number;
}
