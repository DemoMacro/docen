/**
 * @docen/docx — DOCX editor and converter powered by @office-open/docx.
 *
 * @module
 */

// Core: @tiptap/core re-exports + DOCX extension registry
export { docxExtensions, type JSONContent, type AnyExtension } from "./core";

// Editor factory
export { createDocxEditor, type DocxEditorOptions } from "./editor";

// Extensions
export * from "./extensions";

// Converters: DOCX pipeline (DOCX binary ↔ Tiptap JSON)
export {
  parseDOCX,
  generateDOCX,
  generateDOCXSync,
  generateDOCXStream,
  resolveDocument,
  compileDocument,
  DocxManager,
  type DocxGenerateOptions,
} from "./converters/docx";

// Converters: DOCX template patching (placeholder replacement via office-open patchDocument)
export { patchDOCX, type DocxPatchOptions, type DocxPatchContent } from "./converters/patch";

// Converters: Document prepare pipeline (pre-process before compile)
export {
  prepareDocument,
  prepareImages,
  fetchImageHandler,
  type PrepareStep,
  type ImageFetchHandler,
} from "./converters/prepare";

// Converters: HTML pipeline (HTML string ↔ Tiptap JSON)
export { parseHTML, generateHTML } from "./converters/html";

// Converters: Markdown pipeline (Markdown string ↔ Tiptap JSON)
export { parseMarkdown, generateMarkdown } from "./converters/markdown";

// Types
export type * from "./types";
