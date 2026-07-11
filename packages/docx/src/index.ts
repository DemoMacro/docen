/**
 * @docen/docx — DOCX editor and converter powered by @office-open/docx.
 *
 * @module
 */

// Core: @tiptap/core re-exports + DOCX extension registry
export { docxExtensions, type JSONContent, type AnyExtension } from "./core";

// Re-export the OOXML section-properties type (page size/margin/orientation +
// document grid) so the editor layer can type section geometry without a direct
// @office-open/docx dependency.
export type { SectionPropertiesOptions } from "@office-open/docx";
// Re-export the engine's section-geometry defaults (MS Office zh-CN "Normal":
// A4 + top/bottom 1440tw, left/right 1800tw) so editor-side geometry fallbacks
// — content-width for image capping, page measurement — reuse the SAME defaults
// the engine uses to fill an empty sectPr, instead of hardcoding divergent ones.
export { sectionMarginDefaults, sectionPageSizeDefaults } from "@office-open/docx";
// Re-export the engine's length conversion (mm → twips) so the editor layer can
// build OOXML page geometry from mm presets without a direct @office-open/core
// dependency. (@office-open/docx does not re-export this from core.) Sourced from
// the `util` subpath so bundlers can tree-shake the rest of core.
export { convertMillimetersToTwip } from "@office-open/core/util";

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
  normalizeDocument,
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

// Converters: styles → CSS (styles.xml model → scoped editor CSS for rendering)
export {
  stylesToCss,
  quickStyles,
  effectiveRunProps,
  inlineStyles,
  // Style-inheritance primitives shared by the editor's paginator (measure.ts)
  // so pagination measures the SAME effective paragraph properties the renderer
  // emits: a paragraph whose direct attrs are empty still inherits its style's
  // (and its style's basedOn chain's) spacing/indent/run. Without these, measure
  // walked only the direct style → doc-default (skipping basedOn ancestors) and
  // diverged from stylesToCss (which deep-merges the whole chain) → edit≠render
  // page counts on docs with multi-level styles (e.g. Heading1 basedOn Normal).
  indexParagraphStyles,
  defaultParagraphStyleId,
  mergeStyleChain,
  deepMergeInto,
  type StyleEntry,
  type QuickStyleEntry,
  type StylesOptions,
} from "./converters/styles";

// Types
export type * from "./types";
