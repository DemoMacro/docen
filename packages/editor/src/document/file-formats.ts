// Open/save format tables: filename+MIME detection for open(), per-format
// save-picker metadata, the read-only live command set, and the locally
// handled command set (the "wired" basis for ribbon greying).

/** Detect a document's format from its filename + MIME for open(). Extension
 *  first (the picker filters on it), MIME as a fallback for platforms that fill
 *  it in. Throws on an unrecognized type so the caller surfaces the error
 *  rather than silently parsing garbage. */
export function detectOpenFormat(file: File): "docx" | "markdown" {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx")) return "docx";
  if (name.endsWith(".md") || name.endsWith(".markdown")) return "markdown";
  const type = file.type;
  if (type.includes("wordprocessingml.document")) return "docx";
  if (type === "text/markdown") return "markdown";
  throw new Error(`Unsupported file type: ${file.name || type || "(unknown)"}`);
}

/** Per-format metadata for #saveAs: the picker description, the MIME anchoring
 *  its accept filter, and the extension stamped on the suggested name. The MIME
 *  must be a BARE type — showSaveFilePicker rejects accept keys carrying params
 *  (e.g. ";charset=utf-8") with NotSupportedError, so the picker never opens. */
export const SAVE_FORMATS: Record<
  "docx" | "markdown",
  { description: string; mime: string; ext: string }
> = {
  docx: {
    description: "Word Document",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ext: ".docx",
  },
  markdown: { description: "Markdown", mime: "text/markdown", ext: ".md" },
};

/** Commands that stay live when the document is read-only (Viewing mode):
 *  chrome toggles, view panes, the mode switch, save, clipboard reads and
 *  selection — everything else mutates the document and is refused. */
export const READONLY_LIVE: ReadonlySet<string> = new Set([
  "toggle-navigation",
  "zoom",
  "zoom-100",
  "edit-mode",
  "save",
  "copy",
  "select",
  "search",
  "word-count",
  "show-marks",
  "show-comments",
  // View-surface toggles — paint-time view state, not document content.
  "toggle-ruler",
  "toggle-gridlines",
]);

/** Commands handled locally in #onCommand/#onChange (not routed to
 *  editor.commands — they read/write host state the editor can't reach, e.g.
 *  navigation/find/zoom). Together with {@link WIRED_DISPATCH} this is the
 *  "wired" set used to grey out unwired skeleton commands. lang-zh/lang-en are
 *  header menu items, not ribbon commands, so excluded. */
export const LOCAL_HANDLED: ReadonlySet<string> = new Set([
  // #onCommand
  "toggle-navigation",
  "search",
  "replace",
  "page-size",
  "orientation",
  "margins",
  // Columns presets write the current section's w:cols count; Line Numbers
  // toggles the section's w:lnNumType (both via #mutateCurrentSection).
  "columns",
  "line-numbers",
  "zoom",
  "zoom-100",
  "save",
  "insert-picture",
  "show-marks",
  "copy",
  "cut",
  "paste",
  "select",
  "format-painter",
  "edit-mode",
  "word-count",
  // TOC insert/update — dispatch with the bridge's pageOf (page numbers live
  // in the canvas caret map, which editor.commands can't reach).
  "toc",
  "update-toc",
  // Header/footer stories — open through the bridge (the same lifecycle as
  // the band double-click); the Page Number drop seeds a PAGE field.
  "header",
  "footer",
  "page-number",
  // Symbol opens its grid dialog (insertion arrives via symbol:insert);
  // Bookmark prompts for a name and wraps the selection.
  "symbol",
  "bookmark",
  // Paragraph opens the paragraph dialog prefilled from the caret paragraph
  // (the commit arrives via paragraph:ok → the paragraph-dialog-apply
  // command, which stamps every selected paragraph).
  "paragraph-dialog",
  // Footnote prompts for the note text, references the caret and appends the
  // note body to documentExtras.footnotes. Equation drops a placeholder
  // math template (the gallery's fraction/script/radical/sum/integral) at
  // the caret.
  "insert-footnote",
  "equation",
  // Page Color writes the doc-level w:background (doc.attrs.background) from
  // the color-picker's palette value. Page Borders stamps a w:pgBorders
  // preset (none/box/shadow/double/dashed) on the current section. Watermark
  // stamps/removes the preset header shape (every slot, behind-doc).
  // Paragraph Spacing stamps the styles' docDefaults paragraph spacing.
  "page-color",
  "page-border",
  "watermark",
  "paragraph-spacing",
  // View-surface toggles (Word's View → Ruler / Gridlines).
  "toggle-ruler",
  "toggle-gridlines",
  // Link prompts for an address and marks the selection (Word's Insert Link);
  // the context menu's Remove Hyperlink unsets the mark over the clicked link.
  "link",
  "unset-link",
  // New Comment anchors the selection with a Word comment (range markers +
  // a documentExtras.comments entry) — composed in the comments pane, not a
  // prompt; Edit opens the pane (cards edit inline), Delete removes the
  // comment covering the selection, Previous/Next step through the ranges.
  // Show Comments toggles the pane (Word's Review → Show Comments).
  "new-comment",
  "comment",
  "edit-comment",
  "delete-comment",
  "previous-comment",
  "next-comment",
  "show-comments",
  // Text Box / Shapes insert a standalone wps shape run (Shapes reads the
  // gallery preset from the split item's value).
  "text-box",
  "shapes",
  // View → Print Layout runs the same canvas print as the file menu's Print;
  // Outline opens the document-structure pane (Word's outline view maps to
  // the navigation pane here).
  "print-layout",
  "outline",
  // #onChange (data-event)
  "open",
  "save-as",
  "print",
]);
