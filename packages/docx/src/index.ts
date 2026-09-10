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
// Re-export the chart payload type: the chart node's attrs.chart IS a
// ChartOptions verbatim (extensions/chart.ts), and the editor's command layer
// edits it as one — same rationale as SectionPropertiesOptions above. The
// @office-open/docx barrel doesn't re-export the core chart value-domain
// names, so they come straight from @office-open/core (already a dependency).
export type { ChartOptions } from "@office-open/docx";
export type { ChartType, ChartSeriesData, LegendPosition } from "@office-open/core";
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
  parseDOCXSync,
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

// Converters: Markdown pipeline (Markdown string ↔ Tiptap JSON)
export { parseMarkdown, generateMarkdown } from "./converters/markdown";

// Style-facing editor helpers (Quick Styles gallery entries, caret run props)
export {
  quickStyles,
  effectiveRunProps,
  type QuickStyleEntry,
  type StylesOptions,
} from "./converters/styles";

// Style-inheritance primitives (styles.xml model → resolved cascade), shared
// by the layout projection, the CSS route, and the editor's paginator
// (measure.ts) so all of them resolve the SAME effective properties: a
// paragraph whose direct attrs are empty still inherits its style's (and its
// basedOn chain's) spacing/indent/run.
export {
  indexParagraphStyles,
  defaultParagraphStyleId,
  mergeStyleChain,
  deepMergeInto,
  mergeTableStyleProps,
  type StyleEntry,
} from "./style-cascade";

// Types
export type * from "./types";
