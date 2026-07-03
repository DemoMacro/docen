// `docen/docx` subpath — full facade over @docen/docx, the Tiptap DOCX engine.
//
// Re-exports everything: the high-level converters (parseDOCX/generateDOCX,
// parseHTML/generateHTML, parseMarkdown/generateMarkdown), the editor factory
// (createDocxEditor, docxExtensions), the model bridge (resolveDocument /
// compileDocument / prepareDocument), and styles (stylesToCss). Exposed on a
// subpath so the capability layout is symmetric with `docen/editor`; the
// package root re-exports only the high-level converters for convenience.
export * from "@docen/docx";
