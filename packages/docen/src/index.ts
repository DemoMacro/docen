// Pure converter entry — side-effect-free and tree-shakable; re-exports the
// high-level Markdown/HTML/DOCX converters from @docen/docx. The full engine
// (createDocxEditor, docxExtensions, resolve/compile/prepare, styles) is on the
// `docen/docx` subpath, and the web-component editor (<docen-document>) on
// `docen/editor`. Importing converters here never pulls in the engine or UI.
export {
  generateDOCX,
  generateDOCXSync,
  generateDOCXStream,
  parseDOCX,
  patchDOCX,
} from "@docen/docx";
export type { DocxGenerateOptions, DocxPatchContent, DocxPatchOptions } from "@docen/docx";

// HTML / Markdown converters
export { parseHTML, generateHTML, parseMarkdown, generateMarkdown } from "@docen/docx";

// Plain text converters
export { parseText, generateText } from "./text";

// Tiptap core types
export type { JSONContent, Extensions } from "@docen/docx/core";
