export { docxExtensions, tiptapNodeExtensions, tiptapMarkExtensions, DocxKit } from "./extensions";
export type { DocxKitOptions } from "./extensions";
export { Document, createDocument } from "./document";
export { PageBreak } from "./page-break";
// Standalone floating text-box node + its two-element style split, so the
// editor's NodeView can extend WpsShape and render the editable body via a
// contentDOM without re-deriving the engine's EMU/floating geometry.
export { WpsShape } from "./wps-shape";
export { wpsShapeStyles, type WpsShapeStyles, type WpsShapeStandalone } from "./wpg-group";
// Section geometry → CSS mappers, shared by the editor's page node and
// generateHTML so standalone HTML export matches the editor's page geometry.
export {
  resolvePageSize,
  resolveFontName,
  sectionMarginCss,
  sectionLinePitchCss,
  twipsToMm,
} from "./utils";
