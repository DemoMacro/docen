// High-level JSON-in/JSON-out DOCX pipeline (re-exported from @docen/docx)
export { generateDOCX, generateDOCXSync, generateDOCXStream, parseDOCX, patchDOCX } from "./docx";
export type { DocxGenerateOptions, DocxPatchContent, DocxPatchOptions } from "./docx";

// HTML / Markdown converters — signatures match legacy, backed by @docen/docx
export { parseHTML, generateHTML, parseMarkdown, generateMarkdown } from "@docen/docx";

// Plain text converters
export { parseText, generateText } from "./text";

// Tiptap core types (re-exported via @docen/docx/core)
export type { JSONContent, Extensions } from "@docen/docx/core";
