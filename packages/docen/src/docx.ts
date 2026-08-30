// `docen/docx` subpath — full facade over @docen/docx, the Tiptap DOCX engine.
//
// Re-exports everything: the high-level converters (DOCX and Markdown in/out;
// styled HTML input is supported only through the paste pipeline), the editor
// factory (createDocxEditor, docxExtensions), and the model bridge
// (resolveDocument / compileDocument / prepareDocument). Exposed on a
// subpath so the capability layout is symmetric with `docen/editor`; the
// package root re-exports only the high-level converters for convenience.
export * from "@docen/docx";
