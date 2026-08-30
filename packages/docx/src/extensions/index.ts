export * from "./extensions";
// HTML paste input: parsed-HTML body → Tiptap JSON via the schema's rules.
export { parseHTMLBody } from "./paste";
// Section geometry helpers, shared with the editor's page geometry
// (packages/docx layout/project.ts consumes resolvePageSize internally).
// DOCEN_CLIP_MIME + selectionSlicePayload: the docen-lossless clipboard lane
// (a PM slice JSON payload that survives copy/cut → paste with all marks).
export { resolvePageSize, resolveFontName, DOCEN_CLIP_MIME, selectionSlicePayload } from "./utils";
