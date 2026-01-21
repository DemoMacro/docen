// DOCX image information for custom image handler
export interface DocxImageInfo {
  id: string;
  contentType: string;
  data: Uint8Array;
}

// Result of image handling
export interface DocxImageResult {
  src: string;
  alt?: string;
}

// Custom image handler function type for import-docx
export type DocxImageImportHandler = (info: DocxImageInfo) => Promise<DocxImageResult>;

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
