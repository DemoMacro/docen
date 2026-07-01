export * from "./extensions";
// Section geometry → CSS mappers, shared by the editor's page node and
// generateHTML so standalone HTML export matches the editor's page geometry.
export {
  resolvePageSize,
  resolveFontName,
  sectionMarginCss,
  sectionLinePitchCss,
  twipsToMm,
} from "./utils";
