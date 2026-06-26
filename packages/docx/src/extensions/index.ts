export { docxExtensions, tiptapNodeExtensions, tiptapMarkExtensions, DocxKit } from "./extensions";
export type { DocxKitOptions } from "./extensions";
export { Document, createDocument } from "./document";
export { PageBreak } from "./page-break";
// Section geometry → CSS mappers, shared by the editor's page node and
// generateHTML so standalone HTML export matches the editor's page geometry.
export {
  resolvePageSize,
  resolveFontName,
  sectionMarginCss,
  sectionLinePitchCss,
  twipsToMm,
} from "./utils";
