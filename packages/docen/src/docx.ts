// Re-export DOCX functions
export { parseDOCX } from "@docen/import-docx";
export { generateDOCX, patchDOCX } from "@docen/export-docx";

// Re-export types for convenience
export type { DocxImportOptions } from "@docen/import-docx";
export type { DocxExportOptions, DocxPatchOptions, DocxPatchContent } from "@docen/export-docx";
export type * from "@docen/extensions/types";
