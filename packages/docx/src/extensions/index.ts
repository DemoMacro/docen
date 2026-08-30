export * from "./extensions";
// HTML paste input: parsed-HTML body → Tiptap JSON via the schema's rules.
export { parseHTMLBody } from "./paste";
// Section geometry helpers, shared with the editor's page geometry
// (packages/docx layout/project.ts consumes resolvePageSize internally).
export { resolvePageSize, resolveFontName } from "./utils";
