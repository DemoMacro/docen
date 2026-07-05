import {
  convertMillimetersToTwip,
  createDocxEditor,
  effectiveRunProps,
  generateDOCX,
  generateHTML,
  generateMarkdown,
  parseDOCX,
  parseHTML,
  parseMarkdown,
  resolveFontName,
  scrollCaretToTop,
  sectionMarginDefaults,
  sectionPageSizeDefaults,
  stylesToCss,
  twipsToMm,
  type JSONContent,
  type SectionPropertiesOptions,
  type StylesOptions,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { attr, css, customElement, html } from "@microsoft/fast-element";
import type { Mark } from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";
import {
  findNext,
  findPrev,
  getMatchHighlights,
  replaceAll,
  replaceNext,
  setSearchState,
  SearchQuery,
} from "prosemirror-search";

import {
  AddinHost,
  applyTheme,
  mergeRibbonSchema,
  observeLang,
  registerComponents,
  t,
  type DocenAddin,
} from "../ui";
// Side-effect: register the document-specific UI components moved out of the
// shared ui/ barrel — <docen-format-pane> (properties fallback) and
// <docen-outline> (navigation Headings tab).
import "./components/format-pane";
import "./components/outline";
import { createDefaultAddin } from "./addin";
import type { OutlineItem } from "./components/outline";
import { DocenBubbleMenu, defaultBubbleButtons, getBubbleBar } from "./extensions/bubble-menu";
import { WIRED_DISPATCH } from "./extensions/commands";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { clearImageCapCache } from "./extensions/image-cap";
import type { OutlineAnchor } from "./extensions/outline";
import { pageStorageOf } from "./extensions/page-plugin";
import { renderRibbonFromSchema, ribbonActions, ribbonTabs } from "./ribbon";
import { fontNormalRatio } from "./utils/font-metric";
import { clearMeasureCache } from "./utils/measure";
import { unwrapPages, wrapPages } from "./utils/wrap";

/** Escape a host-supplied string for safe interpolation into innerHTML. The
 *  `filename` attribute comes from a user-selected File.name at openDOCX, which
 *  can contain markup — without escaping it flows into #renderHeader's template
 *  and executes. */
const escapeHtml = (s: string): string =>
  s.replace(/[&<>"']/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&#39;",
  );

/** Detect a document's format from its filename + MIME for open(). Extension
 *  first (the picker filters on it), MIME as a fallback for platforms that fill
 *  it in. Throws on an unrecognized type so the caller surfaces the error
 *  rather than silently parsing garbage. */
function detectOpenFormat(file: File): "docx" | "markdown" | "html" {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx")) return "docx";
  if (name.endsWith(".md") || name.endsWith(".markdown")) return "markdown";
  if (name.endsWith(".html") || name.endsWith(".htm")) return "html";
  const type = file.type;
  if (type.includes("wordprocessingml.document")) return "docx";
  if (type === "text/markdown") return "markdown";
  if (type === "text/html") return "html";
  throw new Error(`Unsupported file type: ${file.name || type || "(unknown)"}`);
}

/** Per-format metadata for #saveAs: the picker description, the MIME anchoring
 *  its accept filter, and the extension stamped on the suggested name. The MIME
 *  must be a BARE type — showSaveFilePicker rejects accept keys carrying params
 *  (e.g. ";charset=utf-8") with NotSupportedError, so the picker never opens. */
const SAVE_FORMATS: Record<
  "docx" | "markdown" | "html",
  { description: string; mime: string; ext: string }
> = {
  docx: {
    description: "Word Document",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ext: ".docx",
  },
  markdown: { description: "Markdown", mime: "text/markdown", ext: ".md" },
  html: { description: "HTML Document", mime: "text/html", ext: ".html" },
};

/** Commands handled locally in #onCommand/#onChange (not routed to
 *  editor.commands — they read/write host state the editor can't reach, e.g.
 *  navigation/find/zoom). Together with {@link WIRED_DISPATCH} this is the
 *  "wired" set used to grey out unwired skeleton commands. lang-zh/lang-en are
 *  header menu items, not ribbon commands, so excluded. */
const LOCAL_HANDLED: ReadonlySet<string> = new Set([
  // #onCommand
  "toggle-navigation",
  "search",
  "replace",
  "page-size",
  "orientation",
  "margins",
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
  // #onChange (data-event)
  "open",
  "save-as",
  "print",
]);

const documentStyles = css`
  /* Cascade layers — declared ONCE up front so layer ORDER (not specificity)
       governs priority below. reset strips UA defaults; docxStyles holds the
       document's named styles (styles.xml). Both lose to unlayered rules (page
       geometry, crop marks, inline OOXML styles). */
  @layer reset, docxStyles;
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
  }
  /* Office ribbon group layout helpers — a large button beside stacked rows of
       small icon-only buttons. Applied to light-DOM wrappers in the ribbon. */
  .rb-col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .rb-row {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 2px;
    flex-wrap: wrap;
  }
  /* Small icon-only buttons as a 3-row column-flow grid: buttons stack into
       columns of ≤3 (Word's compact group layout), not a flat single row. */
  .rb-grid {
    display: grid;
    grid-template-rows: repeat(3, auto);
    grid-auto-flow: column;
    gap: 2px;
    align-content: start;
  }
  .rb-vsep {
    width: 1px;
    align-self: stretch;
    background: var(--docen-color-divider, #e1e1e1);
    margin: 0 2px;
  }
  .avatar {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    background: var(--docen-color-brand, #0078d4);
    color: #fff;
    font-size: 10px;
    font-weight: 600;
    margin-inline-end: 4px;
  }
  .avatar-img {
    object-fit: cover;
    background: none;
  }
  /* The editor wrapper (.docen-pages) hosts the Tiptap .ProseMirror, which
       renders one .docen-page NODE per page. The wrapper just centers the
       flow; each page node is its own fixed paper sheet (C-route — see
       CLAUDE.md). */
  .docen-pages .ProseMirror {
    padding: 0;
  }
  /* CSS reset inside the editor surface — strip the browser's UA defaults so
       OOXML properties (inline styles + named styles in the docxStyles layer)
       are the SOLE source of truth. The UA stylesheet gives <p>/<h1-6> a 1em
       margin, resets heading font-size/weight, pads <ul>/<ol>, etc. — all of
       which corrupt pagination (a one-page cover spills to two). These rules
       live in @layer reset, declared BEFORE docxStyles, so the document's named
       styles (.docx-style-*) ALWAYS win regardless of specificity — that is the
       real fix for the Heading bold/centering bug (the old specificity juggling
       is gone). Table borders and list markers are handled in their own rules
       below. */
  @layer reset {
    .docen-pages .ProseMirror p,
    .docen-pages .ProseMirror h1,
    .docen-pages .ProseMirror h2,
    .docen-pages .ProseMirror h3,
    .docen-pages .ProseMirror h4,
    .docen-pages .ProseMirror h5,
    .docen-pages .ProseMirror h6,
    .docen-pages .ProseMirror blockquote,
    .docen-pages .ProseMirror figure,
    .docen-pages .ProseMirror pre,
    .docen-pages .ProseMirror ul,
    .docen-pages .ProseMirror ol {
      margin: 0;
      padding: 0;
    }
    /* Clear the UA heading defaults (2em font-size, bold weight) so headings
         take the doc default unless a named style overrides — same reset layer,
         so a .docx-style-Heading* font-size/font-weight always wins. */
    .docen-page h1,
    .docen-page h2,
    .docen-page h3,
    .docen-page h4,
    .docen-page h5,
    .docen-page h6 {
      font-size: inherit;
      font-weight: inherit;
    }
  }
  /* .ProseMirror's default focus outline paints a black border on every
       click — drop it (the caret + selection still mark focus). */
  .docen-pages .ProseMirror:focus {
    outline: none;
  }
  /* Each page node = a fixed paper sheet. 'height' (NOT min-height) +
       overflow: hidden forces overflow into the next page instead of
       stretching the sheet — the C-route invariant. Geometry comes from
       <docen-document-area> CSS vars inherited through the shadow boundary. */
  .docen-pages .docen-page {
    width: var(--docen-page-width, 210mm);
    height: var(--docen-page-min-height, 297mm);
    overflow: hidden;
    box-sizing: border-box;
    padding: var(--docen-page-margin, 25.4mm);
    margin: 0 auto var(--docen-page-gap, 24px);
    background-color: var(--docen-color-page, #ffffff);
    box-shadow: 0 1px 4px rgba(0, 0, 0, 0.12);
    position: relative;
    /* content-visibility:auto skips layout/paint for off-screen pages — the
         main scroll/load perf win for large (1000+ page) docs. Safe now because
         the paginator no longer reads page DOM for its overflow basis: reflow
         packs against sectionContentDims(section).height (deterministic,
         per-section), and block/row heights are Pretext/model-based (domHeightOf
         only hits hidden passthrough leaves = 0). The page is a fixed-height
         sheet, so per-page contain-intrinsic-size (page-node renderHTML) equals
         the real box and the skipped-page rect matches the laid-out rect → reflow
         still converges. Print forces visible (@media print below); find-in-page,
         selection, and IME keep working — DOM nodes stay, only layout is skipped. */
    content-visibility: auto;
  }
  /* A wrapNone floating drawing (image / wpg group) anchored to its paragraph
       (verticalPosition.relative = paragraph, the OOXML default) renders
       position:absolute. Its offsetParent must be that paragraph, not the page
       box, or top/left resolve from the page top and the drawing floats over
       the heading/body text. Making the anchor <p> position:relative pins the
       drawing to the blank line it belongs on (matches Word: a floating group
       overlays its own empty paragraph, over the body text below it). */
  .docen-pages .docen-page p:has([data-float-anchor="paragraph"]) {
    position: relative;
  }
  /* A paragraph with a leader tab-stop (class docx-tab-leader, set by
       Paragraph.renderHTML from attrs.tabStops) renders a dotted leader between
       the title and page number — MS Word's TOC "....." connector. The tab atom
       (span.docx-tab) flexes to fill the gap and paints the dots; the page
       number sits at the right edge. The paragraph stays single-line (TOC entries
       never wrap), so flex keeps the measured line height intact. */
  .docen-pages .docen-page p.docx-tab-leader {
    display: flex;
    align-items: baseline;
  }
  .docen-pages .docen-page p.docx-tab-leader .docx-tab {
    flex: 1;
    margin: 0 0.3em;
    border-bottom: 1px dotted currentColor;
    opacity: 0.5;
  }
  /* Hyperlinks render via the Link mark (<a href>), but their color / underline
       / cursor come from the run's own rPr (textStyle mark) — NOT the browser's
       default blue / underline / pointer. Reset the UA anchor styling so the
       OOXML run color is the sole source of truth: a TOC entry's runs carry no
       Hyperlink rStyle, so they keep the paragraph style's color (black); a
       styled hyperlink run carries its own color/underline. Word hyperlinks are
       static, so suppress visited/hover/active color shifts too. */
  .docen-pages .docen-page a,
  .docen-pages .docen-page a:visited,
  .docen-pages .docen-page a:hover,
  .docen-pages .docen-page a:active {
    color: inherit;
    text-decoration: inherit;
    cursor: inherit;
  }
  /* Track Changes (w:ins/w:del) — the Insertion/Deletion marks wrap the
       revised text. Word renders inserted text colored + underlined and
       deleted text colored + strikethrough (the text stays visible until
       accept/reject). Colors follow Word's default palette. */
  .docen-pages .docen-page .docen-insertion {
    color: #2e7d32;
    text-decoration: underline;
  }
  .docen-pages .docen-page .docen-deletion {
    color: #c62828;
    text-decoration: line-through;
  }
  /* Images cap to the section content width the way Word caps them: a wider
       image scales DOWN to fit, never upscales. The ImageCap extension sets the
       real width on data-URL images it can sync-decode; this rule is the visual
       fallback for what it skips (http(s) URLs, docs without section geometry)
       so nothing overflows the fixed page box. crop images render an enlarged
       inner <img>, so opt them out or the cap collapses the crop geometry. */
  .docen-pages .docen-page img {
    max-width: 100%;
    height: auto;
  }
  .docen-pages .docen-page span[data-image="crop"] > img {
    max-width: none;
  }
  /* EMF/WMF (Office GDI vector) images can't be decoded by the browser, so
       Image.renderHTML emits a div[data-image=vector] placeholder carrying the
       real art in data-vector-src. Paint it as a dashed, hatched, labeled box
       so the gap reads as "unsupported vector art" instead of a silent empty
       rectangle. Inline width/height (from the image extent) size the box;
       inline-block (set in renderImageStyles) keeps it in-flow. */
  .docen-pages .docen-page [data-image="vector"] {
    box-sizing: border-box;
    border: 1px dashed #b8b8b8;
    background: repeating-linear-gradient(45deg, #f6f6f6, #f6f6f6 9px, #efefef 9px, #efefef 18px);
    color: #9a9a9a;
    font-size: 12px;
    text-align: center;
    padding-top: 6px;
    overflow: hidden;
  }
  /* prosemirror-tables CellSelection tags each selected cell with class
       "selectedCell"; paint it with Word's translucent selection blue so a
       multi-cell selection is actually visible. box-shadow (not
       background-color) overlays a cell's own fill without being beaten by
       the cell's inline background-color. */
  .ProseMirror .selectedCell {
    box-shadow: inset 0 0 0 9999px rgba(0, 120, 215, 0.18);
  }
  /* Crop marks — four L-brackets OUTSIDE the content box, in the page-margin
       gutter, each L's corner pointing AT the editable area: the vertex sits
       just outside a content-box corner and the two 23px legs reach into the
       margin. Drawn on a ::before that covers the whole page (inset: 0) and
       carries the page margin as its own padding, so its content-box == the
       page's content box — that is the origin the background-position
       offsets are measured from (negative for the top legs, positive for the
       bottom), landing each leg in the margin gutter (not over text). The
       ::before is used (not the page node's own background) because a
       ProseMirror-managed node does not paint its own background-image, but its
       pseudo-element does. */
  .docen-pages .docen-page::before {
    content: "";
    position: absolute;
    inset: 0;
    padding: var(--docen-page-margin, 25.4mm);
    pointer-events: none;
    background-image:
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0)),
      linear-gradient(var(--docen-color-crop, #c0c0c0), var(--docen-color-crop, #c0c0c0));
    background-position:
      -24px -2px,
      -2px -24px,
      calc(100% + 24px) -2px,
      calc(100% + 2px) -24px,
      -24px calc(100% + 2px),
      -2px calc(100% + 24px),
      calc(100% + 24px) calc(100% + 2px),
      calc(100% + 2px) calc(100% + 24px);
    background-size:
      23px 1px,
      1px 23px,
      23px 1px,
      1px 23px,
      23px 1px,
      1px 23px,
      23px 1px,
      1px 23px;
    background-origin: content-box;
    background-repeat: no-repeat;
  }
  /* Grey the "Auto-save" label to match its disabled switch (skeleton
       feature), so the label + switch read as one unavailable control, like
       ribbon skeleton buttons. Lifts automatically once the switch loses
       disabled. */
  .autosave-label:has(+ fluent-switch[disabled]) {
    color: var(--docen-color-text-3, #8a8a8a);
  }
  /* While focus is in a ribbon combobox dropdown, the editor is blurred and
       the browser stops painting its selection. The Tiptap Selection extension
       stamps a .selection class on the range so it stays visible. Uses the
       system selection colors (Highlight/HighlightText) to match the browser's
       native ::selection, including high-contrast and custom OS themes. */
  .ProseMirror .selection {
    background: Highlight;
    color: HighlightText;
  }
  /* Placeholder — the Tiptap Placeholder extension only stamps an is-empty
       class + data-placeholder attribute on empty nodes; this CSS paints the
       label. :first-child shows it just on the document's first empty paragraph
       (Word shows the prompt once, at the top); showOnlyCurrent (default) means
       only the caret's empty node carries the attribute. float + height:0 keep
       the label in-flow at the paragraph start without consuming layout, so it
       never pushes real content or skews pagination measurement. */
  .docen-pages .ProseMirror p.is-empty:first-child::before {
    content: attr(data-placeholder);
    color: var(--docen-color-text-3, #8a8a8a);
    pointer-events: none;
    float: left;
    height: 0;
  }
  /* Tables — Word inserts tables in the Table Grid style (a single black
       border on every cell). Without this a freshly inserted table is
       invisible: the docx table extension emits a border only when the node
       carries border attrs, and insertTable creates none. */
  .docen-pages table {
    border-collapse: collapse;
  }
  .docen-pages table td,
  .docen-pages table th {
    border: 1px solid #000;
    /* OOXML defaults w:tcMar top/bottom to 0 (TableNormal); the UA td padding
         (1px) would inflate every row ~2px vs Word. Horizontal padding stays at
         the cell's w:tcMar (set inline by renderTableCellStyles) or 0. */
    padding-block: 0;
  }
  /* Formatting marks (Show/Hide ¶) — Word shows these only while editing
       (non-printing). The show-marks command flips the host [show-marks]
       attribute; the marks themselves live entirely in CSS. */
  /* Pilcrow ¶ is painted by the FormattingMarks extension as a widget
       decoration — CSS ::after on a ProseMirror-managed <p> does not render. */
  /* Zero-width inline-block: the mark hugs the last character and never
       breaks to its own line on a full line. text-indent:0 cancels the
       inherited paragraph indent (an inline-block is a block container), or
       the glyph drifts right. */
  .docen-pages .docen-para-mark {
    color: var(--docen-color-marks, #6e6e6e);
    user-select: none;
    pointer-events: none;
    margin-inline-start: 1px;
    display: inline-block;
    width: 0;
    overflow: visible;
    vertical-align: baseline;
    text-indent: 0;
  }
  /* A page break renders as a Fluent divider with a centered label (Word).
       The NodeView (PageBreakView) supplies the fluent-divider; it is hidden
       unless show-marks is on. */
  .docen-pages [data-type="pageBreak"] {
    display: block;
    line-height: 0;
  }
  .docen-pages [data-type="pageBreak"] fluent-divider {
    display: none;
  }
  :host([show-marks]) .docen-pages [data-type="pageBreak"] {
    margin: 8px 0;
    line-height: normal;
  }
  :host([show-marks]) .docen-pages [data-type="pageBreak"] fluent-divider {
    display: flex;
    font-size: 0.8em;
  }
  /* A section break renders as a Fluent divider after the section-carrying
       paragraph (Word: the boundary only shows the marker while editing). The
       SectionBreakMarks widget supplies the fluent-divider; hidden unless
       show-marks is on — same mechanism as the page-break marker. */
  .docen-pages [data-section-break] {
    line-height: 0;
  }
  .docen-pages [data-section-break] fluent-divider {
    display: none;
  }
  :host([show-marks]) .docen-pages [data-section-break] {
    margin: 8px 0;
    line-height: normal;
  }
  :host([show-marks]) .docen-pages [data-section-break] fluent-divider {
    display: flex;
    font-size: 0.8em;
  }
  /* Find highlights (prosemirror-search) — Word's yellow-match / orange-active. */
  .docen-pages .ProseMirror-search-match {
    background: rgba(255, 235, 59, 0.55);
  }
  .docen-pages .ProseMirror-active-search-match {
    background: rgba(255, 145, 0, 0.75);
  }
  /* Find Results — Office-style match list: each hit rendered with surrounding
       context and a data-from/to for click-to-jump. Padding keeps items off the
       pane edge (the previous "N matches" text butted right against it). */
  .search-results {
    padding: 6px 8px;
    box-sizing: border-box;
  }
  .search-results .result-count {
    font-size: 12px;
    color: var(--docen-color-marks, #6e6e6e);
    padding: 2px 4px 8px;
  }
  .search-results .result-item {
    display: block;
    width: 100%;
    text-align: start;
    border: none;
    background: transparent;
    padding: 5px 8px;
    margin-block-end: 2px;
    border-radius: 4px;
    font-family: inherit;
    font-size: 12px;
    line-height: 1.45;
    color: #3b3b3b;
    cursor: pointer;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .search-results .result-item:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .search-results .result-item mark {
    background: rgba(255, 235, 59, 0.85);
    color: inherit;
    font-weight: 600;
  }
  /* Print: drop shadow/gap/marks; each page node is its own printed sheet. */
  @media print {
    /* @page margin:0 lets the fixed-height page node (297mm) fill the sheet.
         The browser's default ~10mm page margin would otherwise make the page
         taller than the printable area and push it onto a second sheet. The
         page node's own padding (25.4mm) is the document margin, so @page needs
         no margin. size stays auto so non-A4 documents still fit. */
    @page {
      margin: 0;
    }
    .docen-pages .docen-page {
      content-visibility: visible;
      box-shadow: none;
      margin: 0;
      break-after: page;
    }
    /* No trailing blank sheet after the last page. */
    .docen-pages .docen-page:last-child {
      break-after: auto;
    }
    .docen-pages .docen-page::before {
      display: none;
    }
    /* Formatting marks + search highlights never print (editing-only). */
    .docen-pages .docen-para-mark,
    .docen-pages [data-type="pageBreak"],
    .docen-pages [data-section-break],
    .docen-pages .ProseMirror-search-match,
    .docen-pages .ProseMirror-active-search-match {
      display: none !important;
    }
  }
`;

const documentTemplate = html`
  <docen-workspace>
    <docen-title-bar slot="header" part="header"></docen-title-bar>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <docen-task-pane slot="task-pane-start" position="start" part="nav-pane">
      <docen-navigation-pane>
        <docen-outline slot="headings"></docen-outline>
        <div class="search-results" slot="results" part="search-results"></div>
      </docen-navigation-pane>
    </docen-task-pane>
    <docen-document-area>
      <div class="docen-pages" part="page"></div>
    </docen-document-area>
    <docen-task-pane slot="task-pane-end" position="end" part="props-pane">
      <slot name="properties">
        <docen-format-pane></docen-format-pane>
      </slot>
    </docen-task-pane>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <docen-options-dialog part="options"></docen-options-dialog>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <input type="file" id="file-input" accept=".docx,.md,.markdown,.html,.htm" hidden />
  <input type="file" id="image-input" accept="image/*" hidden />
`;

/** Build a nested OutlineItem tree from the flat outline anchor list: each
 *  heading nests under the nearest preceding heading with a smaller level. */
function buildOutlineTree(anchors: readonly OutlineAnchor[]): OutlineItem[] {
  type Node = { id: string; title: string; level: number; children?: Node[] };
  const roots: Node[] = [];
  const stack: Node[] = [];
  for (const a of anchors) {
    const node: Node = { id: a.id, title: a.textContent, level: a.originalLevel };
    while (stack.length && stack[stack.length - 1].level >= a.originalLevel) stack.pop();
    const parent = stack[stack.length - 1];
    if (parent) (parent.children ??= []).push(node);
    else roots.push(node);
    stack.push(node);
  }
  return roots as OutlineItem[];
}

/** prosemirror-search's plugin, wrapped as a Tiptap extension so it loads with
 *  the editor. It stores the active SearchQuery and highlights its matches
 *  (classes ProseMirror-search-match / -active-search-match); the host drives
 *  it via setSearchState and the findNext / findPrev commands. */
/** MS Office standard paper sizes (mm, portrait width × height). Page-setup
 *  presets resolve to raw mm here; <docen-document-area> takes only raw page-width /
 *  page-height, so presets stay in this document layer, not the UI component. */
const PAPER_SIZES: Readonly<Record<string, readonly [number, number]>> = {
  letter: [215.9, 279.4],
  legal: [215.9, 355.6],
  statement: [139.7, 215.9],
  executive: [184.15, 266.7],
  tabloid: [279.4, 431.8],
  a3: [297, 420],
  a4: [210, 297],
  a5: [148, 210],
  a6: [105, 148],
  b5: [182, 257],
};

/** MS Office margin presets (mm, CSS padding list for the page content box:
 *  one value = uniform; two = top/bottom then left/right). `normal` matches the
 *  engine default (@office-open/docx sectionMarginDefaults: top/bottom 25.4mm,
 *  left/right 31.75mm = MS Office zh-CN "Normal") so the canvas fallback and the
 *  document-model sectionProperties agree. */
const MARGINS: Readonly<Record<string, string>> = {
  normal: "25.4mm 31.75mm",
  narrow: "12.7mm",
  moderate: "25.4mm 19.05mm",
  wide: "25.4mm 50.8mm",
};

/** Parse a CSS padding list (mm, 1–4 values) into OOXML page margins (twips),
 *  via the engine's convertMillimetersToTwip (mm → twips). */
function marginTwipsFromCss(css: string): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  const mm = css.split(/\s+/).map((s) => parseFloat(s));
  const [t, r, b, l] =
    mm.length === 1
      ? [mm[0], mm[0], mm[0], mm[0]]
      : mm.length === 2
        ? [mm[0], mm[1], mm[0], mm[1]]
        : [mm[0], mm[1], mm[2] ?? mm[1], mm[3] ?? mm[1]];
  return {
    top: convertMillimetersToTwip(t),
    right: convertMillimetersToTwip(r),
    bottom: convertMillimetersToTwip(b),
    left: convertMillimetersToTwip(l),
  };
}

/** Deep-merge a sectionProperties patch (page.size / page.margin) into a base,
 *  preserving sides/dims the patch omits — so e.g. changing only the margins
 *  keeps the page size. Reuses the engine's SectionPropertiesOptions type. */
function mergeSectionProperties(
  base: SectionPropertiesOptions | null | undefined,
  patch: SectionPropertiesOptions,
): SectionPropertiesOptions {
  const bp = base?.page;
  const pp = patch.page;
  return {
    ...base,
    page: {
      ...bp,
      ...pp,
      size: { ...bp?.size, ...pp?.size },
      margin: { ...bp?.margin, ...pp?.margin },
    },
  };
}

/**
 * `<docen-document>` — a turnkey DOCX editor web component.
 *
 * Wires the Fluent UI host (title-bar + ribbon + document-area) to the `@docen/docx`
 * Tiptap engine, with Pretext-driven offline pagination. Drop it in for an
 * editable, paginated Word surface: the title bar drives file I/O (open/save)
 * and language switching, ribbon commands route to Tiptap, embedded objects open
 * their editors on double-click, and file I/O goes through `parseDOCX`/
 * `generateDOCX`. The title bar + ribbon re-render on locale change.
 */

/**
 * Task pane identifiers, mirroring the Office `<TaskpaneId>` concept. The host
 * ships two built-in panes: `navigation` (start/left) and `properties` (end/right).
 */
export type TaskPaneId = "navigation" | "properties";

/**
 * Visibility mode values, matching `Office.VisibilityMode` (`taskpane` | `hidden`).
 * Carried on {@link docen:taskpane-visibility-change} event details.
 */
export type VisibilityMode = "taskpane" | "hidden";

/** Maps a public {@link TaskPaneId} to the slot position its pane renders in. */
const TASKPANE_POSITION: Record<TaskPaneId, "start" | "end"> = {
  navigation: "start",
  properties: "end",
};

@customElement({ name: "docen-document", template: documentTemplate, styles: documentStyles })
class DocenDocument extends AddinHost<Editor> {
  // ── Reactive attributes (@attr) — the former observedAttributes, re-implemented
  //  as FAST fields. No `reflect` (attribute → property stays one-way). addinsAttr
  //  (attribute "addins") dodges AddinHost.addinsChanged and the `addins` getter.
  @attr editable?: string;
  @attr filename?: string;
  @attr user?: string;
  @attr avatar?: string;
  @attr({ attribute: "section-properties" }) sectionProperties?: string;
  @attr styles?: string;
  @attr({ attribute: "addins" }) addinsAttr?: string;
  @attr theme?: string;

  #editor?: Editor;
  #fileInput?: HTMLInputElement;
  #imageInput?: HTMLInputElement;
  /** Latest TOC anchors, refreshed by TableOfContents.onUpdate; used to resolve
   *  an outline click back to a document position (pos). */
  #anchors: readonly OutlineAnchor[] = [];
  /** Cached doc nodeSize + Office-style word count so caret-move transactions
   *  don't re-walk the whole document (recomputed only when content changes). */
  #lastDocSize = -1;
  #lastWords = 0;
  /** Semantic fingerprint of the last outline tree — id/level/title only. `pos`
   *  shifts on every pagination re-flow but never changes what the pane shows,
   *  so it's excluded; the fingerprint is built from per-anchor arrays (not the
   *  serialized tree) so object key order can never cause a spurious mismatch. */
  #outlineSig = "";
  #unobserveLang?: () => void;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured marks + the pointerup listener that applies them.
  #painterMarks: readonly Mark[] | null = null;
  #painterOff?: () => void;
  /** Current zoom level (percent) applied to the canvas via CSS `zoom`. */
  #zoom = 100;
  /** Debounce timer for the nav-pane search result list — the list rebuilds only
   *  after the user pauses typing (the query dispatches immediately, so Enter /
   *  find-next stays in sync with the last keystroke). */
  #searchTimer?: ReturnType<typeof setTimeout>;
  /** Cached page bounds for the status bar, recomputed only on a doc change so a
   *  caret move (selection-only transaction) reuses them instead of re-walking
   *  every page node each keystroke. */
  #statusDoc?: unknown;
  #pageBounds?: ReadonlyArray<{ offset: number; size: number; section: number }>;

  /** The underlying Tiptap Editor (undefined before connect / after disconnect).
   *  Exposed so a host (the @docen/vue adapter, or any parent element) can drive
   *  commands programmatically — setContent / getJSON / chain / ... — without
   *  routing through the ribbon. */
  get editor(): Editor | undefined {
    return this.#editor;
  }

  /** DocenHost surface — bridge the editor-agnostic `unknown` content contract
   *  to the typed {@link getJSON} / {@link setJSON} API. Addins (and any
   *  DocenHost consumer) read/write content through here without knowing the
   *  runtime is Tiptap JSON. */
  getContent(): unknown {
    return this.getJSON();
  }

  setContent(content: unknown): void {
    if (content && typeof content === "object") {
      this.setJSON(content as JSONContent);
    }
  }

  // ── @attr change callbacks — re-route to the private handlers the old
  //  attributeChangedCallback switch invoked (zero business-logic change). FAST
  //  also fires these during initial attribute hydration; every handler is
  //  guarded (editor/shadowRoot check) so an early fire is a no-op.
  editableChanged(): void {
    this.#editor?.setEditable(this.editable !== "false");
    this.#syncEditModeMenu();
  }

  filenameChanged(): void {
    this.#renderChrome();
  }

  userChanged(): void {
    this.#renderChrome();
  }

  avatarChanged(): void {
    this.#renderChrome();
  }

  sectionPropertiesChanged(): void {
    this.#applySectionPropertiesAttr();
  }

  stylesChanged(): void {
    this.#applyStylesAttr();
  }

  addinsAttrChanged(): void {
    this.#applyAddinsAttr();
  }

  themeChanged(): void {
    this.#applyThemeAttr(this.theme ?? "");
  }

  /** Esc fallback: restore the ribbon to "always shown" after the browser leaves fullscreen. */
  readonly #onFullscreenChange = (): void => {
    if (document.fullscreenElement) return;
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (ribbon?.getAttribute("data-ribbon-mode") === "auto-hide") {
      ribbon.removeAttribute("data-ribbon-mode");
      workspace?.removeAttribute("data-fullscreen");
    }
  };

  /** Status-bar zoom slider → apply the new zoom level. Named (not inline) so it
   *  can be removed on disconnect. */
  readonly #onZoomChange = (event: CustomEvent<{ zoom: number }>): void => {
    this.#setZoom(event.detail.zoom);
  };

  /** Ctrl+= / Ctrl+- / Ctrl+0 zoom, Ctrl+F find (Word behavior). Zoom is
   *  ignored inside ribbon comboboxes and other inputs (so the keystroke reaches
   *  them); Ctrl+F is global. preventDefault blocks the browser's native zoom/find. */
  readonly #onZoomKey = (event: KeyboardEvent): void => {
    // Alt+Q focuses the command search (Office's "Tell me what you want to
    // do" shortcut). Handled before the Ctrl/Meta gate below.
    if (
      event.altKey &&
      !event.ctrlKey &&
      !event.metaKey &&
      (event.key === "q" || event.key === "Q")
    ) {
      event.preventDefault();
      const search = this.shadowRoot?.querySelector("docen-command-search") as HTMLElement | null;
      search?.focus();
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) return;
    // Ctrl+F opens Find, Ctrl+H opens Find & Replace (Word behavior).
    if (event.key === "f" || event.key === "F") {
      event.preventDefault();
      this.#openSearch();
      return;
    }
    if (event.key === "h" || event.key === "H") {
      event.preventDefault();
      this.#openFindReplace();
      return;
    }
    // composedPath()[0] is the real target inside the shadow DOM (e.g. a combobox input).
    const target = event.composedPath()[0] as HTMLElement | null;
    if (target instanceof HTMLElement && target.closest("input, textarea, docen-ribbon-combobox"))
      return;
    const key = event.key;
    if (key === "+" || key === "=") {
      event.preventDefault();
      this.#setZoom(this.#zoom + 10);
    } else if (key === "-" || key === "_") {
      event.preventDefault();
      this.#setZoom(this.#zoom - 10);
    } else if (key === "0") {
      event.preventDefault();
      this.#setZoom(100);
    }
  };

  /** Outline.onUpdate → <docen-outline>. Cache the anchors (so an
   *  outline click resolves to a position) and rebuild the nested tree. */
  #renderOutline(anchors: readonly OutlineAnchor[]): void {
    this.#anchors = anchors;
    const outline = this.shadowRoot?.querySelector("docen-outline");
    if (!outline) return;
    // Fingerprint only what the pane shows (id/level/title). `pos` moves on
    // every pagination re-flow but never changes the outline, so excluding it
    // avoids rebuilding — and flickering — the fluent-tree each pass. Built
    // from per-anchor arrays rather than the serialized tree, so object key
    // order is irrelevant (no dependency on buildOutlineTree's literal field
    // order, unlike a plain JSON.stringify(tree) comparison).
    const sig = anchors
      .map((a) => JSON.stringify([a.id, a.originalLevel, a.textContent]))
      .join("\n");
    if (this.#outlineSig === sig) return;
    this.#outlineSig = sig;
    outline.setAttribute("items", JSON.stringify(buildOutlineTree(anchors)));
  }

  /** Outline click → select the heading at its position and scroll it into view. */
  readonly #onOutlineSelect = (event: CustomEvent<{ id?: string }>): void => {
    const id = event.detail?.id;
    const editor = this.#editor;
    if (!id || !editor) return;
    const anchor = this.#anchors.find((a) => a.id === id);
    if (!anchor) return;
    editor.chain().focus().setTextSelection(anchor.pos).run();
    scrollCaretToTop(editor.view);
  };

  /** navigation:search → set the active query; matches highlight live. */
  readonly #onSearch = (event: CustomEvent<{ query?: string }>): void => {
    const editor = this.#editor;
    if (!editor) return;
    const query = new SearchQuery({ search: event.detail?.query ?? "", caseSensitive: false });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    // Debounce the result-list rebuild (O(matches) DOM nodes per keystroke); the
    // query already dispatched above, so find-next reads the live search state.
    clearTimeout(this.#searchTimer);
    this.#searchTimer = setTimeout(() => this.#updateSearchResults(), 120);
  };

  /** ribbon-mode-change → drive browser fullscreen + status-bar hide.
   *  auto-hide = Full Screen (Office); any other mode exits it. Named so it can
   *  be removed on disconnect (an anonymous listener would leak on reconnect). */
  readonly #onRibbonModeChange = (event: Event): void => {
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (!workspace) return;
    const mode = (event as CustomEvent<{ mode: string }>).detail.mode;
    if (mode === "auto-hide") {
      void this.requestFullscreen?.().catch(() => {});
      workspace.setAttribute("data-fullscreen", "");
    } else {
      if (document.fullscreenElement) void document.exitFullscreen?.().catch(() => {});
      workspace.removeAttribute("data-fullscreen");
    }
  };

  /** navigation:find → jump to the next/previous match (prosemirror-search). */
  readonly #onFind = (event: CustomEvent<{ direction: "next" | "prev" }>): void => {
    const editor = this.#editor;
    if (!editor) return;
    (event.detail.direction === "prev" ? findPrev : findNext)(editor.state, editor.view.dispatch);
  };

  /** Stamp the Results slot with the live match list — each hit rendered with
   *  surrounding context and a data-from/to for click-to-jump (Word's Results
   *  pane lists every match with context, not just a count). */
  #updateSearchResults(): void {
    const editor = this.#editor;
    const slot = this.shadowRoot?.querySelector(".search-results");
    if (!slot) return;
    const decos = editor ? getMatchHighlights(editor.state).find() : [];
    slot.replaceChildren();
    const header = document.createElement("div");
    header.className = "result-count";
    header.textContent =
      decos.length > 0
        ? `${decos.length} ${t("search.matches", this)}`
        : t("search.noResults", this);
    slot.append(header);
    if (!editor || decos.length === 0) return;
    const doc = editor.state.doc;
    const RADIUS = 24;
    for (const deco of decos) {
      const { from, to } = deco as { from: number; to: number };
      const before = doc.textBetween(Math.max(0, from - RADIUS), from, " ");
      const after = doc.textBetween(to, Math.min(doc.content.size, to + RADIUS), " ");
      const item = document.createElement("button");
      item.type = "button";
      item.className = "result-item";
      item.dataset.from = String(from);
      item.dataset.to = String(to);
      if (before) {
        const span = document.createElement("span");
        span.textContent = `…${before}`;
        item.append(span);
      }
      const hit = document.createElement("mark");
      hit.textContent = doc.textBetween(from, to, " ");
      item.append(hit);
      if (after) {
        const span = document.createElement("span");
        span.textContent = `${after}…`;
        item.append(span);
      }
      slot.append(item);
    }
  }

  /** Click a Results entry → select that match range and scroll it into view. */
  readonly #onSearchResultClick = (event: Event): void => {
    const editor = this.#editor;
    if (!editor) return;
    const item = (event.target as HTMLElement | null)?.closest(".result-item");
    if (!(item instanceof HTMLElement)) return;
    const from = Number(item.dataset.from);
    const to = Number(item.dataset.to);
    if (!Number.isFinite(from) || !Number.isFinite(to)) return;
    editor.chain().focus().setTextSelection({ from, to }).run();
    scrollCaretToTop(editor.view);
  };

  /** Ctrl+F → open the nav pane and focus its search box (Word behavior). */
  #openSearch(): void {
    const taskPane = this.shadowRoot?.querySelector('docen-task-pane[position="start"]') as
      | (HTMLElement & { open: boolean })
      | null;
    if (taskPane) taskPane.open = true;
    const input = this.shadowRoot
      ?.querySelector("docen-navigation-pane")
      ?.shadowRoot?.querySelector("[part='search-input']") as
      | (HTMLElement & { select: () => void })
      | null;
    input?.focus();
    input?.select?.();
  }

  /** Ctrl+H / ribbon Replace → open the Find & Replace dialog. */
  #openFindReplace(): void {
    const dialog = this.shadowRoot?.querySelector("docen-find-replace-dialog") as
      | (HTMLElement & { show: () => void })
      | null;
    dialog?.show();
  }

  /** find-replace:action → drive prosemirror-search (query highlights, find-next,
   *  replace-next = replace + advance, replace-all). Each action re-stamps the
   *  query so Find/Replace/options are always current. */
  readonly #onFindReplace = (
    event: CustomEvent<{
      action: string;
      find: string;
      replace: string;
      caseSensitive: boolean;
      wholeWord: boolean;
    }>,
  ): void => {
    const editor = this.#editor;
    if (!editor) return;
    const { action, find, replace, caseSensitive, wholeWord } = event.detail ?? {};
    const query = new SearchQuery({ search: find, replace, caseSensitive, wholeWord });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    if (action === "find-next") findNext(editor.state, editor.view.dispatch);
    else if (action === "replace-next") replaceNext(editor.state, editor.view.dispatch);
    else if (action === "replace-all") replaceAll(editor.state, editor.view.dispatch);
  };

  /** Paste from the system clipboard as plain text. navigator.clipboard is the
   *  reliable path; execCommand("paste") is the fallback (often blocked). */
  async #paste(): Promise<void> {
    const editor = this.#editor;
    if (!editor) return;
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      editor.commands.focus();
      document.execCommand("paste");
      return;
    }
    if (text) editor.chain().focus().insertContent(text).run();
  }

  /** Editing → Select menu. "all" uses the official selectAll() command (an
   *  AllSelection that crosses page isolating boundaries); "objects"/"similar"
   *  are placeholders. */
  #select(value?: string): void {
    const editor = this.#editor;
    if (!editor) return;
    if ((value ?? "all") !== "all") return;
    editor.chain().focus().selectAll().run();
  }

  /** Editing → Find drop-down → Go To: prompt for a page number and move the
   *  caret to that page, scrolling it into view. */
  #goToPage(): void {
    const editor = this.#editor;
    if (!editor) return;
    const input = window.prompt(t("ribbon.opt.go-to-prompt", this));
    if (input == null) return;
    const page = parseInt(input, 10);
    if (!Number.isFinite(page) || page < 1) return;
    let count = 0;
    let target = -1;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "page") {
        count++;
        if (count === page) {
          target = pos + 1;
          return false;
        }
      }
      return true;
    });
    if (target < 0) return;
    editor.chain().focus().setTextSelection(target).run();
    scrollCaretToTop(editor.view);
  }

  /** Format Painter: on first click, capture the current selection's marks and
   *  arm a one-shot pointerup listener; the next non-empty selection receives
   *  those marks and disarms the painter. A second click cancels. */
  #toggleFormatPainter(): void {
    if (this.#painterMarks) {
      this.#stopFormatPainter();
      return;
    }
    const editor = this.#editor;
    if (!editor || editor.state.selection.empty) return;
    this.#painterMarks = editor.state.selection.$from.marks();
    this.toggleAttribute("format-painter", true);
    const dom = editor.view.dom;
    const onUp = (): void => {
      const ed = this.#editor;
      if (!ed) return;
      const { from, to, empty } = ed.state.selection;
      if (!empty && this.#painterMarks) {
        const tr = ed.state.tr;
        for (const mark of this.#painterMarks) tr.addMark(from, to, mark);
        ed.view.dispatch(tr);
      }
      this.#stopFormatPainter();
    };
    dom.addEventListener("pointerup", onUp, { once: true });
    this.#painterOff = () => dom.removeEventListener("pointerup", onUp);
  }

  #stopFormatPainter(): void {
    this.#painterMarks = null;
    this.removeAttribute("format-painter");
    this.#painterOff?.();
    this.#painterOff = undefined;
  }

  /** Mirror the font name / size and paragraph style at the caret into the
   *  ribbon comboboxes — Word behavior: the boxes report the formatting at the
   *  cursor, not a fixed default. Re-runs on every editor transaction (caret
   *  moves, marks change). Font/size read the browser-resolved computed style
   *  so the full style inheritance chain (direct run props → paragraph-style →
   *  document defaults) is reflected. */
  #setupFontSync(): void {
    const editor = this.#editor;
    if (!editor) return;
    const sync = (): void => {
      this.#syncFontControls();
      this.#syncStyleControl();
      this.#updateStatus();
    };
    editor.on("transaction", sync);
    sync();
    this.#fontSyncCleanup = (): void => {
      editor.off("transaction", sync);
    };
  }

  #syncFontControls(): void {
    const editor = this.#editor;
    if (!editor) return;
    const fontDisplay = this.#effectiveFontAt(editor);
    const sizeDisplay = this.#effectiveSizeAt(editor);
    const fontCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-name"]',
    );
    const sizeCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-size"]',
    );
    if (fontCb && fontCb.getAttribute("value") !== fontDisplay) {
      fontCb.setAttribute("value", fontDisplay);
    }
    if (sizeCb && sizeCb.getAttribute("value") !== sizeDisplay) {
      sizeCb.setAttribute("value", sizeDisplay);
    }
  }

  /** The font name at the caret, resolved through the style inheritance chain
   *  (direct run props → paragraph style → basedOn → document defaults) in the
   *  document's own units — no px conversion. */
  #effectiveFontAt(editor: Editor): string {
    const direct = editor.getAttributes("textStyle");
    const { font } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      direct,
    );
    return font ?? "";
  }

  /** The font size at the caret in points (Word's unit), resolved through the
   *  style inheritance chain — no px conversion. */
  #effectiveSizeAt(editor: Editor): string {
    const direct = editor.getAttributes("textStyle");
    const { size } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      direct,
    );
    return size != null ? String(size) : "";
  }

  /** The loaded document's styles model (doc.attrs.styles), or null. */
  #docStyles(editor: Editor): StylesOptions | null {
    return (editor.state.doc.attrs?.styles as StylesOptions | undefined) ?? null;
  }

  /** The paragraph-style id at the caret (HeadingLevel literal for headings,
   *  pStyle id for paragraphs). */
  #currentStyleId(editor: Editor): string | null {
    const h = editor.getAttributes("heading") as { styleId?: unknown };
    const p = editor.getAttributes("paragraph") as { styleId?: unknown };
    if (typeof h?.styleId === "string" && h.styleId) return h.styleId;
    if (typeof p?.styleId === "string" && p.styleId) return p.styleId;
    return null;
  }

  /** Mirror the paragraph style at the caret into the Styles gallery combobox —
   *  its value is the current paragraph's styleId (the HeadingLevel literal for
   *  headings, the pStyle id for paragraphs, or "Normal" when the paragraph
   *  carries none). The combobox matches the value against its gallery items to
   *  show the style's display name. */
  #syncStyleControl(): void {
    const editor = this.#editor;
    if (!editor) return;
    const headingAttrs = editor.getAttributes("heading") as { styleId?: unknown };
    const paraAttrs = editor.getAttributes("paragraph") as { styleId?: unknown };
    const styleId =
      typeof headingAttrs?.styleId === "string" && headingAttrs.styleId
        ? headingAttrs.styleId
        : typeof paraAttrs?.styleId === "string" && paraAttrs.styleId
          ? paraAttrs.styleId
          : "";
    const value = styleId || "Normal";
    const cb = this.shadowRoot?.querySelector<HTMLElement>('docen-ribbon-combobox[event="style"]');
    if (cb && cb.getAttribute("value") !== value) cb.setAttribute("value", value);
  }

  async connectedCallback(): Promise<void> {
    super.connectedCallback();
    await registerComponents();
    applyTheme(this.getAttribute("theme") === "dark" ? "dark" : "light");

    this.#fileInput = this.shadowRoot!.querySelector<HTMLInputElement>("#file-input")!;
    this.#imageInput = this.shadowRoot!.querySelector<HTMLInputElement>("#image-input")!;
    this.#renderChrome();
    // Once attributes: initial task-pane visibility (Office `setStartupBehavior`
    // equivalent). Absent → closed; present → open. Read once on connect —
    // runtime toggles go through showTaskpane/hideTaskpane.
    this.#setTaskpane("navigation", this.hasAttribute("navigation-pane"));
    this.#setTaskpane("properties", this.hasAttribute("properties-pane"));
    // Once attribute: initial zoom level (percent). Runtime zoom goes through
    // setZoom / the status-bar slider / Ctrl+ -/=/0.
    const initialZoom = this.getAttribute("zoom");
    if (initialZoom) this.#setZoom(Number(initialZoom) || 100);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.addEventListener("zoom:change", this.#onZoomChange as EventListener);

    const page = this.shadowRoot!.querySelector<HTMLDivElement>(".docen-pages");
    if (!page) return;

    // Fonts must be loaded before pagination measures, else Pretext drifts
    // from the browser's actual line layout (see rendering-engine-choices).
    await document.fonts?.ready;

    // Wrap the initial content into doc > page+ (the editing schema). The
    // page node never enters DOCX — wrapPages/unwrapPages bridge it at the
    // editor layer, so DOCX round-trip stays transparent.
    const contentAttr = this.getAttribute("content");
    // Declarative section-properties / styles (JSON) seed doc-level attrs so a
    // host can bootstrap page setup + named styles without openDOCX/setJSON.
    const initAttrs = this.#readInitAttrs();
    const baseDoc = wrapPages(contentAttr ? parseHTML(contentAttr) : undefined);
    const initialDoc =
      Object.keys(initAttrs).length > 0
        ? { ...baseDoc, attrs: { ...baseDoc.attrs, ...initAttrs } }
        : baseDoc;
    // The default document add-in contributes the engine extensions + every
    // wired ribbon command. Registered before the editor mounts so its
    // extensions seed the schema. Ribbon events route straight to the engine
    // via DocumentCommands (editor.chain().<event>), not addin.commands.
    const defaultAddin = createDefaultAddin({
      onOutlineUpdate: (anchors) => this.#renderOutline(anchors),
    });
    this.addAddin(defaultAddin);
    // Declarative external add-ins (JSON `addins` attribute) register after the
    // default so their ribbon tabs append to the built-ins via mergeRibbonSchema.
    this.#applyAddinsAttr();

    // Bubble-menu buttons: built-in defaults + addin contributions, merged at
    // boot — symmetric to `ribbonTabs(styles) + mergeRibbonSchema(addins)` in
    // #renderChrome. The bar extension stays OUT of defaultAddin.extensions
    // (the host owns the merge, like the ribbon), so it's configured here with
    // the assembled list. Runtime `addAddin({ bubbleMenu })` re-merges in
    // addinsChanged (the bar's buttons are @observable), no re-mount needed.
    const bubbleButtons = [...defaultBubbleButtons(), ...this.mergedBubbleMenu()];
    this.#editor = createDocxEditor({
      element: page,
      content: initialDoc,
      // Spellcheck defaults OFF — Chromium's spellcheck is a major perf cost on
      // large documents (ProseMirror community-confirmed). Opt in via the
      // Review ribbon's spell-check button (spellcheck="true" attribute).
      spellcheck: this.getAttribute("spellcheck") === "true",
      editable: this.getAttribute("editable") !== "false",
      // Engine extensions come from the default add-in (see addin.ts); the
      // bubble-menu extension is layered on with the host-merged buttons.
      extensions: [
        ...(defaultAddin.extensions ?? []),
        DocenBubbleMenu.configure({ commands: bubbleButtons }),
      ],
    });
    this.#applyDocStyles();

    // Mirror the caret's font/size into the ribbon comboboxes (Word behavior).
    this.#setupFontSync();
    // Stamp the status bar (page count / caret page / zoom) once laid out.
    this.#updateStatus();

    // Default page setup (Word defaults): A4 portrait + Normal margins. The
    // canvas already defaults to 210×297; apply the margin preset so the
    // content box matches Word and pagination measures correctly. Skip when the
    // host declared `section-properties` — that already seeded
    // doc.attrs.sectionProperties via the initial doc, so just sync the canvas.
    if (this.hasAttribute("section-properties")) {
      this.#syncCanvasFromSection();
    } else {
      this.#setMargins("normal");
    }

    // command = ribbon buttons; change = menu items + auto-save switch. Listen
    // on the shadow root so non-composed Fluent events (menu-item "change")
    // reach us, not just composed ones (ribbon "command").
    this.shadowRoot!.addEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot!.addEventListener("change", this.#onChange as EventListener);
    this.#fileInput.addEventListener("change", this.#onFileChange);
    this.#imageInput.addEventListener("change", this.#onImageChange);
    // Outline (Headings tab) → jump to the clicked heading.
    this.shadowRoot!.querySelector("docen-outline")?.addEventListener(
      "outline:select",
      this.#onOutlineSelect as EventListener,
    );
    // Nav-pane search → Find (live highlight, next/prev, results count).
    this.addEventListener("navigation:search", this.#onSearch as EventListener);
    this.addEventListener("navigation:find", this.#onFind as EventListener);
    // Click a Results entry → jump to that match (delegated on the container).
    this.shadowRoot!.querySelector(".search-results")?.addEventListener(
      "click",
      this.#onSearchResultClick as EventListener,
    );
    // Find & Replace dialog → Replace / Replace All (prosemirror-search).
    this.shadowRoot!.querySelector("docen-find-replace-dialog")?.addEventListener(
      "find-replace:action",
      this.#onFindReplace as EventListener,
    );
    // Options dialog — ok; status-bar language indicator — lang:change.
    this.shadowRoot!.querySelector("docen-options-dialog")?.addEventListener(
      "options:ok",
      this.#onOptionsOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "lang:change",
      this.#onLangChange as EventListener,
    );

    // Re-render header + ribbon when the page locale (<html lang>) changes.
    this.#unobserveLang = observeLang(() => this.#renderChrome());

    // Ribbon Display Options → drive browser fullscreen + status-bar hide.
    // auto-hide = Full Screen (Office); any other mode exits it.
    const ribbon = this.shadowRoot!.querySelector("docen-ribbon");
    ribbon?.addEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    // Emit docen:change on every content change (autosave driver) and docen:ready
    // once the editor is live — both bubble out so a host can react.
    this.#editor?.on("transaction", this.#onTransaction);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
    this.dispatchEvent(new CustomEvent("docen:ready", { bubbles: true, composed: true }));
  }

  disconnectedCallback(): void {
    this.shadowRoot?.removeEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot?.removeEventListener("change", this.#onChange as EventListener);
    this.#fileInput?.removeEventListener("change", this.#onFileChange);
    this.#imageInput?.removeEventListener("change", this.#onImageChange);
    this.shadowRoot
      ?.querySelector("docen-outline")
      ?.removeEventListener("outline:select", this.#onOutlineSelect as EventListener);
    this.removeEventListener("navigation:search", this.#onSearch as EventListener);
    this.removeEventListener("navigation:find", this.#onFind as EventListener);
    this.shadowRoot
      ?.querySelector(".search-results")
      ?.removeEventListener("click", this.#onSearchResultClick as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#onFindReplace as EventListener);
    this.shadowRoot
      ?.querySelector("docen-options-dialog")
      ?.removeEventListener("options:ok", this.#onOptionsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("lang:change", this.#onLangChange as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("zoom:change", this.#onZoomChange as EventListener);
    this.#unobserveLang?.();
    this.#editor?.off("transaction", this.#onTransaction);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.shadowRoot
      ?.querySelector("docen-ribbon")
      ?.removeEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    clearTimeout(this.#searchTimer);
    this.#fontSyncCleanup?.();
    this.#editor?.destroy();
    super.disconnectedCallback();
  }

  /** App header markup — i18n labels plus the host-supplied `user` / `filename`
   *  (not translated: identity and the file name come from the app via
   *  attributes, not the locale table). The auto-save label sits to the left of
   *  its toggle (Word layout), so the switch carries only an aria-label. */
  #renderHeader(): string {
    const user = this.getAttribute("user") ?? "";
    const avatar = this.getAttribute("avatar") ?? "";
    const filename = this.getAttribute("filename") ?? t("header.doc-name");
    const initial = user.trim().charAt(0).toUpperCase();
    const avatarMarkup = avatar
      ? `<img class="avatar avatar-img" src="${escapeHtml(avatar)}" alt="" />`
      : initial
        ? `<span class="avatar">${initial}</span>`
        : "";
    const autosave = t("header.autosave");
    return `
          <div slot="start" style="display:flex;align-items:center;gap:4px">
            <span style="font-weight:600;font-size:13px;padding-inline:6px">${t("header.brand")}</span>
            <span class="autosave-label">${autosave}</span>
            <!-- auto-save is skeleton-only — disabled (greyed, non-interactive)
                 until the feature is wired; the autosave case in onChange is a
                 no-op. Remove disabled (and re-add checked) when it lands. -->
            <fluent-switch data-event="autosave" disabled aria-label="${autosave}"></fluent-switch>
            <docen-ribbon-button icon="save" label="${t("header.save")}" event="save" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="undo" label="${t("header.undo")}" event="undo" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="redo" label="${t("header.redo")}" event="redo" icon-only></docen-ribbon-button>
            <fluent-menu>
              <fluent-menu-button slot="trigger" appearance="subtle">${escapeHtml(filename)}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new">${t("header.new")}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="open">${t("header.open")}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="save-as">${t("header.save-as")}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-markdown">${t("header.save-as-markdown")}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-html">${t("header.save-as-html")}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="print">${t("header.print")}</fluent-menu-item>
                <fluent-menu-item data-event="options">${t("header.options")}</fluent-menu-item>
              </fluent-menu-list>
            </fluent-menu>
          </div>
          <docen-command-search slot="search"></docen-command-search>
          <div slot="end" style="display:flex;align-items:center;gap:4px">
            <span style="display:inline-flex;align-items:center;gap:6px;padding-inline:6px">${avatarMarkup}${escapeHtml(user)}</span>
          </div>`;
  }

  /** Stamp the header + ribbon markup for the active locale (re-run on lang change). */
  #renderChrome(): void {
    const root = this.shadowRoot;
    // FAST fires @attr change callbacks during element upgrade, BEFORE the
    // template is stamped (connectedCallback runs after) — the shadowRoot
    // exists but is empty, so the title-bar query is null. Bail until stamped;
    // connectedCallback's explicit call does the first render.
    const titleBar = root?.querySelector("docen-title-bar");
    if (!root || !titleBar) return;
    const styles = this.#editor?.state.doc.attrs?.styles ?? null;
    titleBar.innerHTML = this.#renderHeader();
    // Built-in tabs (Home/Insert/… with the live style gallery) come from
    // ribbonTabs; external add-ins layer their own tabs on top via
    // mergeRibbonSchema. The default add-in contributes no ribbon, so without
    // extra add-ins this is just the built-in set.
    const tabs = [...ribbonTabs(styles), ...mergeRibbonSchema(this.addins)];
    const ribbonEl = root.querySelector("docen-ribbon")!;
    ribbonEl.replaceChildren(renderRibbonFromSchema(tabs, ribbonActions()));
    // Feed the full ribbon schema (built-in tabs + addin contributions) to the
    // command search so it can flatten and index every command. Re-runs on
    // lang/addin change since #renderChrome is the single chrome re-stamp.
    const searchEl = root.querySelector("docen-command-search") as
      | (HTMLElement & { setTabs(tabs: readonly unknown[]): void })
      | null;
    searchEl?.setTabs(tabs);
    this.#applyRibbonGreying();
    this.#syncEditModeMenu();
    this.#renderPanes();
  }

  /** Addin registry changed (add-in registered/removed) — re-stamp the ribbon
   *  so an external add-in's ribbon contribution appears. The default add-in
   *  contributes no ribbon, so this is a no-op for it; only extra add-ins add
   *  tabs. */
  protected addinsChanged(): void {
    this.#renderChrome();
    // Re-merge the bubble-menu buttons so a runtime addAddin's `bubbleMenu`
    // takes effect immediately — symmetric to the ribbon re-render above. The
    // bar's `commands` is @observable, so re-assignment re-renders the row and
    // re-injects icons without rebuilding the BubbleMenu plugin. No-op before
    // the editor boots (bar is null until addProseMirrorPlugins runs).
    const bar = getBubbleBar();
    if (bar) {
      bar.commands = [...defaultBubbleButtons(), ...this.mergedBubbleMenu()];
    }
  }

  /** Stamp pane titles + status text for the active locale (re-run on lang change). */
  #renderPanes(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const navPane = root.querySelector('docen-task-pane[position="start"]');
    if (navPane) navPane.setAttribute("title", t("pane.navigation"));
    const propsPane = root.querySelector('docen-task-pane[position="end"]');
    if (propsPane) propsPane.setAttribute("title", t("pane.properties"));
    // Page-break / section-break divider labels live in the editor view
    // (NodeView / widget decoration), not the ribbon — update them alongside
    // the chrome so a locale change relabels them.
    root.querySelectorAll<HTMLElement>("fluent-divider[data-pb]").forEach((d) => {
      d.textContent = t("ribbon.cmd.page-break");
    });
    root.querySelectorAll<HTMLElement>("fluent-divider[data-sb]").forEach((d) => {
      d.textContent = t("ribbon.cmd.section-break");
    });
    // Status bar is dynamic (page count / caret page / zoom) — re-stamp it so a
    // locale change re-localizes the text too.
    this.#updateStatus();
    // The placeholder prompt comes from a closure over t(); its decoration set
    // is only rebuilt on a view update, so dispatch a meta-only transaction to
    // force a rebuild against the now-current locale (empty doc → new prompt).
    const view = this.#editor?.view;
    if (view) view.dispatch(view.state.tr.setMeta("docen-i18n", true));
  }

  /** Add-in ids currently registered from the `addins` attribute. Tracked so
   *  editing the attribute at runtime removes add-ins that fell out (addAddin
   *  alone is idempotent on add but can't detect a deletion). */
  #addinAttrIds = new Set<string>();

  /** Sync external add-ins with the `addins` JSON attribute: register new ids,
   *  remove ids no longer present. JSON can't carry functions, so only ribbon
   *  data contributions cross this boundary; command handlers stay in JS
   *  (addAddin with a full object). */
  #applyAddinsAttr(): void {
    const raw = this.getAttribute("addins");
    const next = new Set<string>();
    if (raw) {
      let parsed: unknown;
      try {
        parsed = JSON.parse(raw);
      } catch {
        return;
      }
      if (Array.isArray(parsed)) {
        for (const item of parsed) {
          if (
            item &&
            typeof item === "object" &&
            typeof (item as { id?: unknown }).id === "string"
          ) {
            const id = (item as { id: string }).id;
            next.add(id);
            if (!this.#addinAttrIds.has(id)) this.addAddin(item as DocenAddin<this>);
          }
        }
      }
    }
    // Remove add-ins that fell out of the attribute (covers editing it to drop
    // a tab at runtime, or removing the attribute entirely).
    for (const id of this.#addinAttrIds) {
      if (!next.has(id)) this.removeAddin(id);
    }
    this.#addinAttrIds = next;
  }

  /** Apply the `theme` attribute: switch the Fluent theme (light/dark). */
  #applyThemeAttr(value: string): void {
    applyTheme(value === "dark" ? "dark" : "light");
  }

  /** Grey out ribbon commands that have no handler (skeleton buttons). Runs
   *  after every ribbon re-stamp; fresh elements start un-disabled, so this is
   *  the single place `disabled` is applied. Only controls that support
   *  `disabled` (button/split-button/toggle-button) are greyed — combobox /
   *  color-picker lack it and live in wired tabs anyway. */
  #applyRibbonGreying(): void {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (!ribbon) return;
    const wired = this.#wiredCommands();
    ribbon
      .querySelectorAll<HTMLElement>(
        "docen-ribbon-button[event], docen-ribbon-split-button[event], docen-ribbon-toggle-button[event]",
      )
      .forEach((el) => {
        const event = el.getAttribute("event");
        if (event && !wired.has(event)) el.setAttribute("disabled", "");
      });
  }

  /** Re-stamp the tab-row "Editing" menu so its label + checked item match the
   *  editor's live editable state (initial render, after a switch, and on
   *  locale change — #renderChrome re-stamps the ribbon, so this runs after
   *  #applyRibbonGreying to override the static default items). */
  #syncEditModeMenu(): void {
    const menu = this.shadowRoot?.querySelector('docen-ribbon-menu[event="edit-mode"]');
    if (!menu) return;
    const editable = this.#editor?.isEditable ?? true;
    menu.setAttribute("label", t(editable ? "ribbon.opt.editing" : "ribbon.opt.viewing"));
    menu.setAttribute(
      "items",
      JSON.stringify([
        { text: t("ribbon.opt.editing"), event: "edit-mode", value: "edit", checked: editable },
        { text: t("ribbon.opt.viewing"), event: "edit-mode", value: "view", checked: !editable },
      ]),
    );
  }

  /** The full set of wired command names (Tiptap dispatch + locally handled +
   *  addin commands). External add-ins register non-Tiptap actions (e.g. open a
   *  URL) via `commands`; their keys count as wired so {@link #applyRibbonGreying}
   *  doesn't disable the controls that dispatch them. */
  #wiredCommands(): Set<string> {
    const wired = new Set<string>([...WIRED_DISPATCH, ...LOCAL_HANDLED]);
    for (const addin of this.addins) {
      if (!addin.commands) continue;
      for (const key of Object.keys(addin.commands)) wired.add(key);
    }
    return wired;
  }

  /** Dispatch a cancelable event; returns true when a host preventDefaulted it
   *  (i.e. took over the action). Lets save/open/print/new work out-of-box yet
   *  stay overridable. */
  #emitCancelable(
    name: "docen:save" | "docen:save-as" | "docen:open" | "docen:new" | "docen:print",
    detail?: { format?: "docx" | "markdown" | "html" },
  ): boolean {
    const event = new CustomEvent(name, {
      bubbles: true,
      composed: true,
      cancelable: true,
      detail,
    });
    this.dispatchEvent(event);
    return event.defaultPrevented;
  }

  /** docen:change — fired on every doc-changing transaction (autosave driver,
   *  mirroring OnlyOffice's onDocumentStateChange). Selection-only transactions
   *  are skipped. */
  readonly #onTransaction = (props: { transaction: { docChanged: boolean } }): void => {
    if (props.transaction.docChanged) {
      this.dispatchEvent(
        new CustomEvent("docen:change", { bubbles: true, composed: true, detail: { dirty: true } }),
      );
    }
  };

  /** Toggle a task pane open/closed (ribbon View → toggle-navigation). */
  #togglePane(id: TaskPaneId): void {
    this.#setTaskpane(id, !this.getTaskpaneState(id));
  }

  /** Apply a paper-size preset (a4/letter/…) to the canvas as raw mm, then
   *  re-paginate. Also writes the size into the document-model sectionProperties
   *  (Word stores page setup in the sectPr) so render/measure/image-cap/export
   *  share one geometry source; the canvas attrs are now the rendering fallback
   *  + the zoom surface's page-width source. */
  #setPageSize(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    const size = value ? PAPER_SIZES[value] : undefined;
    if (canvas && size) {
      canvas.setAttribute("page-width", String(size[0]));
      canvas.setAttribute("page-height", String(size[1]));
    }
    if (size) {
      this.#updateSectionGeometry({
        page: {
          size: {
            width: convertMillimetersToTwip(size[0]),
            height: convertMillimetersToTwip(size[1]),
          },
        },
      });
    }
    this.#refreshGeometry();
  }

  /** Apply orientation (portrait/landscape) to the canvas, then re-paginate.
   *  portrait clears the attribute (the canvas default); landscape sets it and
   *  the canvas swaps width/min-height via :host([orientation]). Also writes
   *  orientation onto page.size, deep-merged with the current (or engine-default)
   *  size so resolvePageSize can swap edges for landscape. */
  #setOrientation(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    if (canvas && value) {
      if (value === "landscape") canvas.setAttribute("orientation", "landscape");
      else canvas.removeAttribute("orientation");
    }
    if (value) {
      const cur = (
        this.#editor?.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions }
      )?.sectionProperties?.page?.size;
      const size =
        cur && typeof cur.width === "number" && typeof cur.height === "number"
          ? cur
          : { width: sectionPageSizeDefaults.WIDTH, height: sectionPageSizeDefaults.HEIGHT };
      this.#updateSectionGeometry({
        page: { size: { ...size, orientation: value as "portrait" | "landscape" } },
      });
    }
    this.#refreshGeometry();
  }

  /** Apply a margin preset (normal/narrow/…) to the canvas as a CSS padding
   *  list, then re-paginate. Also writes the margins into the document-model
   *  sectionProperties so a page-setup change actually re-caps images and
   *  re-renders (the canvas CSS alone wouldn't, once a sectPr is inlined). */
  #setMargins(value?: string): void {
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    if (canvas && value && MARGINS[value]) canvas.setAttribute("margin", MARGINS[value]);
    if (value && MARGINS[value]) {
      this.#updateSectionGeometry({ page: { margin: marginTwipsFromCss(MARGINS[value]) } });
    }
    this.#refreshGeometry();
  }

  /** Deep-merge a sectionProperties patch into the CURRENT section's sectPr and
   *  dispatch it — Word's "this section" semantics. The current section is the
   *  one holding the caret: its sectPr rides on its last paragraph (the first
   *  section-carrying paragraph at/after the caret), or, when the caret is in the
   *  final section, on doc.attrs.sectionProperties (the body-level sectPr).
   *  Reflow re-stamps each page's geometry from its section's sectPr, so an edit
   *  only affects the caret's section — multi-section docs keep per-section page
   *  setups, and render/measure/image-cap/export all see the change. */
  #updateSectionGeometry(patch: SectionPropertiesOptions): void {
    const editor = this.#editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const from = editor.state.selection.from;
    // First section-carrying paragraph at/after the caret = the current
    // section's last paragraph (OOXML: its sectPr ends that section).
    let targetPos: number | null = null;
    doc.descendants((node, nodePos) => {
      if (targetPos != null || nodePos < from) return true;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        targetPos = nodePos;
        return false;
      }
      return true;
    });
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        const cur = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        tr.setNodeMarkup(targetPos, undefined, {
          ...node.attrs,
          sectionProperties: mergeSectionProperties(cur, patch),
        });
      }
    } else {
      // Caret in the final section (no section-carrying paragraph at/after it) —
      // its sectPr is body-level (doc.attrs.sectionProperties).
      const cur = (doc.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
      tr.setDocAttribute("sectionProperties", mergeSectionProperties(cur, patch));
    }
    editor.view.dispatch(tr);
  }

  /** Re-paginate after a page-size / orientation / margin change. RAF so the
   *  new canvas geometry applies first; PagePlugin re-reads the page height
   *  from the DOM and re-solves the breaks. */
  #refreshGeometry(): void {
    requestAnimationFrame(() => {
      if (this.#editor) pageStorageOf(this.#editor).repaginate();
    });
  }

  /** Parse the declarative `section-properties` / `styles` attributes (JSON).
   *  Lets a host bootstrap page setup + named styles without openDOCX/setJSON.
   *  Malformed JSON is ignored (warned) so a typo never breaks the editor. */
  #readInitAttrs(): {
    sectionProperties?: SectionPropertiesOptions;
    styles?: StylesOptions;
  } {
    const out: { sectionProperties?: SectionPropertiesOptions; styles?: StylesOptions } = {};
    const sp = this.getAttribute("section-properties");
    if (sp) {
      try {
        out.sectionProperties = JSON.parse(sp) as SectionPropertiesOptions;
      } catch {
        console.warn("[docen-document] invalid section-properties JSON — ignored");
      }
    }
    const st = this.getAttribute("styles");
    if (st) {
      try {
        out.styles = JSON.parse(st) as StylesOptions;
      } catch {
        console.warn("[docen-document] invalid styles JSON — ignored");
      }
    }
    return out;
  }

  /** Mirror doc.attrs.sectionProperties onto the canvas attributes (page-width /
   *  page-height / margin / orientation) so zoom-to-page-width and the CSS
   *  fallback stay in sync, then re-paginate. Sides absent in the model fall
   *  back to the engine default margins. */
  #syncCanvasFromSection(): void {
    const editor = this.#editor;
    if (!editor) return;
    const sp = (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
    const page = sp?.page;
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    if (canvas && page) {
      const size = page.size;
      if (size && typeof size.width === "number" && typeof size.height === "number") {
        canvas.setAttribute("page-width", twipsToMm(size.width));
        canvas.setAttribute("page-height", twipsToMm(size.height));
      }
      if (size?.orientation === "landscape") canvas.setAttribute("orientation", "landscape");
      else canvas.removeAttribute("orientation");
      const m = page.margin;
      const def = sectionMarginDefaults;
      const mm = (v: unknown, d: number): string => twipsToMm(typeof v === "number" ? v : d);
      canvas.setAttribute(
        "margin",
        `${mm(m?.top, def.TOP)} ${mm(m?.right, def.RIGHT)} ${mm(m?.bottom, def.BOTTOM)} ${mm(m?.left, def.LEFT)}`,
      );
    }
    this.#refreshGeometry();
  }

  /** Runtime `section-properties` change: deep-merge into the body section's
   *  sectPr (a default doc is single-section) and re-sync the canvas. */
  #applySectionPropertiesAttr(): void {
    const editor = this.#editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().sectionProperties;
    if (!parsed) return;
    const cur = (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("sectionProperties", mergeSectionProperties(cur, parsed)),
    );
    this.#syncCanvasFromSection();
  }

  /** Runtime `styles` change: replace doc.attrs.styles and re-inject the CSS. */
  #applyStylesAttr(): void {
    const editor = this.#editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().styles;
    if (parsed === undefined) return;
    editor.view.dispatch(editor.state.tr.setDocAttribute("styles", parsed));
    this.#applyDocStyles();
  }

  /** Apply a zoom level (percent, clamped 10–500) to the canvas and refresh the
   *  status bar. CSS `zoom` rescales the pages and reflows the scroll surface.
   *  Idempotent (no-op on no change) and dispatches `docen:zoom-change` on a
   *  real flip — so the host, status-bar slider, and external listeners stay in
   *  sync through one funnel (Office `Office.Document.zoom.set` equivalent). */
  #setZoom(pct: number): void {
    const next = Math.max(10, Math.min(500, Math.round(pct)));
    if (next === this.#zoom) return;
    this.#zoom = next;
    this.shadowRoot?.querySelector("docen-document-area")?.setAttribute("zoom", String(this.#zoom));
    this.#updateStatus();
    this.dispatchEvent(
      new CustomEvent("docen:zoom-change", {
        bubbles: true,
        composed: true,
        detail: { zoom: this.#zoom },
      }),
    );
  }

  /** Resolve a ribbon zoom preset to a percent. Numeric presets ("200", "100",
   *  "75", "50") map directly to that zoom level; "page-width" scales the page
   *  to fill the canvas width (mm → px @96dpi). */
  #zoomPreset(preset: string): void {
    if (/^\d+$/.test(preset)) return this.#setZoom(Number(preset));
    if (preset !== "page-width") return;
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    if (!canvas) return;
    const MM2PX = 96 / 25.4;
    const pw = parseFloat(canvas.getAttribute("page-width") ?? "210") * MM2PX;
    this.#setZoom((canvas.clientWidth / pw) * 100);
  }

  /** Refresh the status bar to mirror Word's bottom row: the left cluster is
   *  the caret's section, then "Page X of Y", then the word count; the right
   *  cluster is the zoom slider value + percent. Runs on every transaction
   *  (caret moves, a re-flow changes the page count) and on zoom / locale change.
   *
   *  The section number is O(pages): each page node carries its section's
   *  sectionProperties attrs (one shared reference across a section's pages),
   *  so the count increments only when that reference changes — no paragraph
   *  walk. The word count is cached by doc nodeSize so caret moves skip
   *  re-walking the full document (CharacterCount.words() regexes all text). */
  #updateStatus(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const bar = root.querySelector<HTMLElement>("docen-status-bar");
    const editor = this.#editor;
    let page = 0;
    let total = 0;
    let section = 1;
    if (editor) {
      const doc = editor.state.doc;
      const from = editor.state.selection.from;
      // Cache page bounds across selection-only transactions (caret moves);
      // recompute only when the doc changed. Avoids a full doc.forEach + per-page
      // attrs read on every keystroke/caret move on large multi-page documents.
      if (doc !== this.#statusDoc) {
        const bounds: { offset: number; size: number; section: number }[] = [];
        let prevSp: unknown;
        let sectionCount = 0;
        let firstPage = true;
        doc.forEach((node, offset) => {
          if (node.type.name !== "page") return;
          const sp = (node.attrs as { sectionProperties?: unknown }).sectionProperties ?? null;
          // A section's pages share one sectionProperties reference, so the count
          // rises only at a real section boundary (or on the very first page).
          if (firstPage || sp !== prevSp) sectionCount++;
          firstPage = false;
          prevSp = sp;
          bounds.push({ offset, size: node.nodeSize, section: sectionCount });
        });
        this.#pageBounds = bounds;
        this.#statusDoc = doc;
      }
      const bounds = this.#pageBounds!;
      total = bounds.length;
      for (let i = 0; i < bounds.length; i++) {
        const b = bounds[i];
        if (from > b.offset && from <= b.offset + b.size) {
          page = i + 1;
          section = b.section;
          break;
        }
      }
    }
    // Word count is cached by doc nodeSize so caret moves skip re-walking the
    // full document (CharacterCount.words() regexes all text).
    const docSize = editor?.state.doc.nodeSize ?? 0;
    if (docSize !== this.#lastDocSize) {
      const cc = editor?.storage.characterCount as { words?: () => number } | undefined;
      this.#lastWords = cc?.words?.() ?? 0;
      this.#lastDocSize = docSize;
    }
    // Push the numeric state to <docen-status-bar>; it localizes + renders.
    if (bar) {
      bar.setAttribute("section", String(section));
      bar.setAttribute("page", String(page || 1));
      bar.setAttribute("total", String(total || 1));
      bar.setAttribute("words", String(this.#lastWords));
      bar.setAttribute("zoom", String(this.#zoom));
    }
  }

  readonly #onCommand = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const { event: name, value } = event.detail ?? {};
    if (typeof name !== "string") return;
    // UI chrome actions are handled locally and need no Tiptap editor.
    if (name === "toggle-navigation") {
      this.#togglePane("navigation");
      return;
    }
    // Find (ribbon Home → Editing → Find, or Ctrl+F) → open the nav-pane search.
    if (name === "search") {
      // Find drop-down → Go To jumps to a page; the main button and Find
      // open the nav-pane search box.
      if (value === "go-to") this.#goToPage();
      else this.#openSearch();
      return;
    }
    // Replace (ribbon Home → Editing → Replace, or Ctrl+H) → Find & Replace dialog.
    if (name === "replace") {
      this.#openFindReplace();
      return;
    }
    // Page setup actions are handled locally (they change the canvas/page, not
    // the Tiptap doc) and need no editor command.
    if (name === "page-size") {
      this.#setPageSize(value);
      return;
    }
    if (name === "orientation") {
      this.#setOrientation(value);
      return;
    }
    if (name === "margins") {
      this.#setMargins(value);
      return;
    }
    // Zoom is a canvas action (not a Tiptap command): step in, or apply a
    // preset from the split menu (200/100/75/50/page-width); the split's
    // main button sets 100%.
    if (name === "zoom") {
      this.#setZoom(this.#zoom + 10);
      return;
    }
    if (name === "zoom-100") {
      if (value) this.#zoomPreset(value);
      else this.#setZoom(100);
      return;
    }
    if (!this.#editor) return;
    // Edit / View mode — toggle the editor's editable state (tab-row "Editing"
    // menu); then re-stamp the menu so its label + checked item follow.
    if (name === "edit-mode") {
      this.#editor.setEditable(value !== "view");
      this.#syncEditModeMenu();
      return;
    }
    // "save" is a document action, not a Tiptap command — handle locally,
    // unless the host took over via docen:save (preventDefault).
    if (name === "save") {
      if (!this.#emitCancelable("docen:save")) void this.#saveAs();
      return;
    }
    // Picture needs a file picker — open it, then insert the chosen image.
    if (name === "insert-picture") {
      this.#imageInput?.click();
      return;
    }
    // Formatting marks toggle — FormattingMarks paints the paragraph mark via a
    // widget decoration; the host [show-marks] attr drives the page-break
    // divider CSS.
    if (name === "show-marks") {
      this.setShowMarks(!this.getShowMarks());
      return;
    }
    // Clipboard — execCommand copy/cut acts on the editor's DOM selection;
    // paste reads the system clipboard (contenteditable execCommand paste is
    // blocked in most browsers).
    if (name === "copy" || name === "cut") {
      this.#editor.commands.focus();
      document.execCommand(name);
      return;
    }
    if (name === "paste") {
      void this.#paste();
      return;
    }
    // Editing → Select: selectAll() spans every page (bypassing page
    // isolating); objects/similar are not yet wired.
    if (name === "select") {
      this.#select(value);
      return;
    }
    // Format Painter — toggle capture/apply of the current run's marks.
    if (name === "format-painter") {
      this.#toggleFormatPainter();
      return;
    }
    // Built-in commands route to editor.chain().focus().<event>(value).run() —
    // DocumentCommands registers every ribbon event as a native Tiptap command.
    // A user add-in overrides one by contributing a Tiptap extension whose
    // addCommands redefines the same name (Tiptap's native override mechanism).
    const editor = this.#editor;
    if (editor) {
      const chain = editor.chain().focus() as unknown as Record<
        string,
        (value?: string) => { run: () => void }
      >;
      const cmd = chain[name];
      if (typeof cmd === "function") {
        cmd(value).run();
        return;
      }
    }
    // Not a Tiptap command — route to the first add-in that declares it. This
    // covers non-Tiptap actions contributed by external add-ins (e.g. a Help
    // button that opens a URL) that Tiptap can't express.
    this.dispatchCommand(name, value);
  };

  /** Menu items and the auto-save switch carry their action in `data-event`. */
  readonly #onChange = (event: Event): void => {
    const name = (event.target as HTMLElement)?.dataset?.event;
    if (!name) return;
    switch (name) {
      case "open":
        // Host can take over via docen:open (preventDefault); else open the
        // picker — #onFileChange auto-detects docx/md/html from the extension.
        if (!this.#emitCancelable("docen:open")) this.#pickFile();
        break;
      case "save-as":
        if (!this.#emitCancelable("docen:save-as", { format: "docx" })) void this.#saveAs("docx");
        break;
      case "save-as-markdown":
        if (!this.#emitCancelable("docen:save-as", { format: "markdown" }))
          void this.#saveAs("markdown");
        break;
      case "save-as-html":
        if (!this.#emitCancelable("docen:save-as", { format: "html" })) void this.#saveAs("html");
        break;
      case "print":
        if (!this.#emitCancelable("docen:print")) this.#print();
        break;
      case "new":
        // No built-in "new" — always hand to the host (docen:new).
        this.#emitCancelable("docen:new");
        break;
      case "options": {
        // Filename menu → open the Options dialog (UI language).
        const optionsEl = this.shadowRoot?.querySelector("docen-options-dialog");
        if (optionsEl) {
          optionsEl.setAttribute("locale", document.documentElement.lang || "zh-CN");
          (optionsEl as unknown as { show?: () => void }).show?.();
        }
        break;
      }
      // autosave: skeleton — wired when that feature lands.
    }
  };

  readonly #onLangChange = (event: Event): void => {
    const lang = (event as CustomEvent<{ lang: string }>).detail?.lang;
    if (lang) document.documentElement.lang = lang;
  };

  /** Options dialog 确定 — commit the UI language. */
  readonly #onOptionsOk = (event: Event): void => {
    const lang = (event as CustomEvent<{ lang?: string }>).detail?.lang;
    if (lang && document.documentElement.lang !== lang) {
      document.documentElement.lang = lang;
    }
  };

  /** Open the OS file picker. The accept filter on the input element covers
   *  .docx/.md/.markdown/.html/.htm; #onFileChange routes the chosen file by
   *  extension via open(). */
  #pickFile(): void {
    this.#fileInput?.click();
  }

  readonly #onFileChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    // Reset so picking the same file twice still fires `change`.
    input.value = "";
    if (!file) return;
    void this.open(file);
  };

  /** Insert the picked image as a data URL. Width/height are left unset — the
   *  browser shows natural size, and prepareImages fills them on DOCX export. */
  readonly #onImageChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      this.#editor
        ?.chain()
        .focus()
        .insertContent({ type: "image", attrs: { src: reader.result } })
        .run();
    };
    reader.readAsDataURL(file);
  };

  /** Save the document in the given format via the native Save As dialog
   *  (showSaveFilePicker) when available so the user picks the location and name;
   *  falls back to a plain download otherwise. The header filename is updated to
   *  match the saved name. HTML is wrapped in a full document so the file
   *  renders standalone in a browser. */
  async #saveAs(format: "docx" | "markdown" | "html" = "docx"): Promise<void> {
    const cfg = SAVE_FORMATS[format];
    // saveDOCX returns a buffer; the text formats return a string (HTML wrapped
    // for standalone rendering).
    const data =
      format === "docx"
        ? await this.saveDOCX()
        : format === "markdown"
          ? this.saveMarkdown()
          : this.#wrapHtmlDocument(this.saveHTML());
    const blob = new Blob([data as BlobPart], { type: cfg.mime });
    // Re-stamp the extension so a .docx opened then saved as Markdown does not
    // keep its .docx name.
    const baseName = (this.getAttribute("filename")?.trim() || t("header.doc-name")).replace(
      /\.(docx|md|markdown|htm|html|txt)$/i,
      "",
    );
    const suggestedName = baseName + cfg.ext;
    const picker = (
      window as unknown as {
        showSaveFilePicker?: (opts: {
          suggestedName?: string;
          types?: Array<{ description?: string; accept: Record<string, string[]> }>;
        }) => Promise<{
          name: string;
          createWritable: () => Promise<{
            write: (data: Blob | BufferSource | string) => Promise<void>;
            close: () => Promise<void>;
          }>;
        }>;
      }
    ).showSaveFilePicker;
    if (picker) {
      try {
        const handle = await picker({
          suggestedName,
          types: [{ description: cfg.description, accept: { [cfg.mime]: [cfg.ext] } }],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        this.setAttribute("filename", handle.name);
        this.#renderChrome();
        return;
      } catch {
        // The user cancelled the picker (AbortError) or it was blocked — do NOT
        // fall back to a download, which would save despite the cancel. The
        // download fallback below only covers browsers without the picker.
        return;
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = suggestedName;
    a.click();
    URL.revokeObjectURL(url);
  }

  /** Wrap a generated HTML body fragment in a full document so a saved .html
   *  file renders standalone — generateHTML returns <section> fragments only. */
  #wrapHtmlDocument(body: string): string {
    const title = (this.getAttribute("filename")?.trim() || t("header.doc-name")).replace(
      /\.[^.]+$/,
      "",
    );
    return `<!DOCTYPE html><html lang="${escapeHtml(document.documentElement.lang || "en")}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(title)}</title></head><body>${body}</body></html>`;
  }

  /** Print the document: window.print() — the @media print rules in this
   *  template, the workspace, and the canvas hide the chrome and render only
   *  the page content. */
  #print(): void {
    window.print();
  }

  /** Common load path for openDOCX/openMarkdown/openHTML: inject the doc styles,
   *  adopt a filename, replace the whole doc node, then re-paginate once layout
   *  has settled. Markdown/HTML carry no doc-level styles, so `json.attrs` has
   *  none and #injectDocStyles clears the CSS. */
  #applyOpenedJSON(json: JSONContent, filename?: string): void {
    const editor = this.#editor;
    if (!editor) return;
    // Inject the styles CSS BEFORE rendering so the first paint carries the
    // document's real fonts/sizes (see setJSON for the rationale).
    this.#injectDocStyles((json.attrs as { styles?: StylesOptions } | undefined)?.styles);
    // Set filename before #applyDocStyles so its single #renderChrome reflects
    // both the Styles gallery and the filename header — previously this was a
    // second #renderChrome call duplicating the one inside #applyDocStyles.
    if (filename) this.setAttribute("filename", filename);
    // Replace the whole doc node (content + doc-level attrs) — see #loadDoc.
    this.#loadDoc(editor, wrapPages(json));
    this.#applyDocStyles();
    // Paginate once the new document has laid out. setContent renders the DOM
    // synchronously, but the browser commits layout on the NEXT frame —
    // measuring in the same frame reads half-laid-out blocks and the breaks
    // come out wrong (the first page overflows, so the pages look uneven).
    this.#repaginateAfterLoad(editor);
  }

  /** Load a file into the editor, auto-detecting its format from the extension
   *  (.docx → DOCX, .md/.markdown → Markdown, .html/.htm → HTML). This is the
   *  single entry point the filename-menu "Open…" uses; openDOCX/openMarkdown/
   *  openHTML remain for when the caller already knows the format (e.g. loading
   *  a server-fetched docx buffer that has no filename). Throws on an
   *  unrecognized extension. */
  async open(file: File): Promise<void> {
    const format = detectOpenFormat(file);
    if (format === "docx") return this.openDOCX(file);
    if (format === "markdown") return this.openMarkdown(file);
    return this.openHTML(file);
  }

  /** Load a .docx into the editor from a File or a buffer (ArrayBuffer /
   *  Uint8Array). A File also adopts its name as the filename; a bare buffer
   *  carries no name. parseDOCX is synchronous, but this is async so a File's
   *  bytes can be awaited. */
  async openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void> {
    const buffer = input instanceof File ? await input.arrayBuffer() : input;
    this.#applyOpenedJSON(parseDOCX(buffer), input instanceof File ? input.name : undefined);
  }

  /** Load a Markdown file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. */
  async openMarkdown(input: File | string): Promise<void> {
    const text = typeof input === "string" ? input : await input.text();
    this.#applyOpenedJSON(parseMarkdown(text), typeof input === "string" ? undefined : input.name);
  }

  /** Load an HTML file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. Section geometry and the page
   *  background are doc-level metadata that round-trip via DOCX, not HTML, so
   *  only the content is restored. */
  async openHTML(input: File | string): Promise<void> {
    const text = typeof input === "string" ? input : await input.text();
    this.#applyOpenedJSON(parseHTML(text), typeof input === "string" ? undefined : input.name);
  }

  /** Re-paginate after loading a document, once its layout has settled.
   *
   *  Two passes: the first rAF runs as soon as the browser has committed layout
   *  for the just-set content (correct for text and tables — the common case,
   *  and fast enough to read as "instant"). Fonts and images load after that
   *  frame and change block heights, so a second pass re-measures once they're
   *  ready. No debounce: an import must paginate within a frame, not after
   *  300ms — the old deferred pass lingered on stale seams for seconds. */
  #repaginateAfterLoad(editor: Editor): void {
    const run = (): void => {
      if (this.#editor === editor && !editor.isDestroyed) pageStorageOf(editor).repaginate();
    };
    requestAnimationFrame(run);
    // Second pass once fonts are ready: fonts loading after the first pass
    // change canvas measureText widths, so the Pretext prepare cache (keyed on
    // text+font+letterSpacing) is stale — clear it before re-measuring with the
    // real fonts. Images no longer gate this pass: image paragraphs measure from
    // node.attrs (not the <img> DOM), so a still-loading image can't change page
    // breaks — awaiting it only stalled the second pass (a large doc's many
    // images stalled it near-indefinitely).
    const fonts = document.fonts?.ready ?? Promise.resolve();
    void fonts.then(() => {
      clearMeasureCache();
      requestAnimationFrame(run);
    });
  }

  /** Serialize the current document to a DOCX buffer. */
  async saveDOCX(): Promise<Uint8Array> {
    const json = unwrapPages(this.#editor!.getJSON() as JSONContent);
    const buffer = await generateDOCX(json);
    return buffer as unknown as Uint8Array;
  }

  /** Serialize the current document to a Markdown string. */
  saveMarkdown(): string {
    return generateMarkdown(unwrapPages(this.#editor!.getJSON() as JSONContent));
  }

  /** Serialize the current document to an HTML body fragment (no
   *  <html>/<!DOCTYPE> wrapper — #saveAs wraps it for a standalone file). */
  saveHTML(): string {
    return generateHTML(unwrapPages(this.#editor!.getJSON() as JSONContent));
  }

  /** Current document as Tiptap JSON. */
  getJSON(): JSONContent {
    // Unwrap page nodes — external consumers see flat doc > block+ (page
    // nodes are editor-only; transparent in the public API).
    const editor = this.#editor;
    return editor ? unwrapPages(editor.getJSON()) : ({} as JSONContent);
  }

  /** Replace the document with Tiptap JSON. */
  setJSON(json: JSONContent): void {
    const editor = this.#editor;
    if (!editor) return;
    // Inject the styles CSS BEFORE rendering so the first paint already carries
    // the document's real fonts/sizes — without this, a large doc's synchronous
    // load + image decode can defer the <style> repaint and the page renders
    // unstyled. Read styles from the incoming JSON (editor.state is still the
    // old doc here; wrapPages preserves doc-level attrs).
    this.#injectDocStyles((json.attrs as { styles?: StylesOptions } | undefined)?.styles);
    // Wrap on the way in — external callers pass flat doc > block+. Tiptap's
    // setContent only swaps content (it drops doc-level attrs like styles/core),
    // so replace the whole doc node via a fresh EditorState to preserve them.
    this.#loadDoc(editor, wrapPages(json));
    this.#applyDocStyles();
    // Paginate once layout has settled — parity with openDOCX (setContent renders
    // synchronously, but the browser commits layout on the next frame).
    this.#repaginateAfterLoad(editor);
  }

  /** Replace the whole doc node (content + doc-level attrs) via a fresh
   *  EditorState. Tiptap's setContent only swaps content and drops doc-level
   *  attrs; this carries them (styles/core/sectionProperties). updateState
   *  bypasses appendTransaction/onTransaction, so extensions that react to doc
   *  changes wouldn't wake — dispatch a docChanged tr (re-stamp the first
   *  block's attrs, a no-op visually) to trigger them: TableOfContents injects
   *  heading ids + fires onUpdate, PagePlugin schedules a re-flow. */
  #loadDoc(editor: Editor, doc: JSONContent): void {
    // Reset per-document image caches so a prior doc's decoded sizes / failed
    // fetches neither leak nor suppress a legit re-fetch in the new document.
    clearImageCapCache();
    editor.view.updateState(
      EditorState.create({ doc: editor.schema.nodeFromJSON(doc), plugins: editor.state.plugins }),
    );
    if (editor.isDestroyed) return;
    // updateState bypasses appendTransaction, so extensions that react to doc
    // changes wouldn't wake. Dispatch a docChanged tr to fire them: TableOfContents
    // injects heading ids + emits onUpdate, PagePlugin schedules a re-flow. The
    // tr re-stamps the LAST leaf block's OWN attrs — a true no-op (same node,
    // same attrs) — so nothing is clobbered. Targeting the last leaf (not pos 1)
    // sidesteps the page node's `isolating` boundary entirely: the old
    // setNodeMarkup(1) resolved INTO the page's first child and could overwrite
    // its attrs, which is how the first heading lost its styleId (and with it
    // its Heading1 bold/centering) on load.
    const state = editor.state;
    // Collect leaf blocks; the last one (deepest, rightmost) is the re-stamp
    // target. A const array (mutated, not reassigned) keeps TS's control-flow
    // analysis happy — a `let` reassigned inside the callback reads as `never`
    // outside it (TS can't see the callback runs synchronously).
    const leaves: { pos: number; attrs: Record<string, unknown> }[] = [];
    state.doc.nodesBetween(0, state.doc.content.size, (node, pos) => {
      if (node.isText) return;
      if (node.isTextblock || node.isLeaf) {
        leaves.push({ pos, attrs: node.attrs as Record<string, unknown> });
      }
      // Don't descend into textblocks (their text isn't a markup target).
      return node.isTextblock ? false : undefined;
    });
    const last = leaves[leaves.length - 1];
    if (last) {
      editor.view.dispatch(state.tr.setNodeMarkup(last.pos, undefined, last.attrs));
    }
  }

  /** Inject the document's named styles (styles.xml) as scoped CSS so imported
   *  headings/body text render with their real font/size/color instead of the
   *  browser default. Called after every content load (create / import / setJSON). */
  /** Inject the document's named-styles CSS (styles.xml → scoped <style>) into
   *  the shadow root. Idempotent — reuses the #docen-doc-styles element across
   *  calls. Called BOTH before #loadDoc (by openDOCX/setJSON, so the first paint
   *  is already styled) and inside #applyDocStyles (refresh after doc attrs
   *  settle). Split out so loaders can inject before rendering. */
  #injectDocStyles(styles: StylesOptions | null | undefined): void {
    const root = this.shadowRoot;
    if (!root) return;
    const css = stylesToCss(styles, ".docen-page");
    const existing = root.querySelector("#docen-doc-styles");
    if (css) {
      const styleEl = (existing ?? document.createElement("style")) as HTMLStyleElement;
      styleEl.id = "docen-doc-styles";
      // Wrap in @layer docxStyles so these named styles beat the reset layer
      // (layer order, not specificity) yet stay below unlayered inline styles.
      styleEl.textContent = "@layer docxStyles {\n" + css + "\n}";
      if (!existing) root.append(styleEl);
    } else {
      existing?.remove();
    }
  }

  /** Apply the loaded document's styles + chrome + geometry. Called after every
   *  content load (create / import / setJSON). The styles <style> is also
   *  injected BEFORE #loadDoc by the loaders (first paint styled); this refreshes
   *  it from the now-settled doc attrs, re-stamps the ribbon (Styles gallery),
   *  and applies font-metric + section geometry. */
  #applyDocStyles(): void {
    const editor = this.#editor;
    if (!editor) return;
    this.#injectDocStyles(editor.state.doc.attrs?.styles);
    // Re-stamp the ribbon so the Styles gallery reflects the loaded document's
    // style library (named + custom paragraph styles) — see styleItems().
    this.#renderChrome();
    // Font metric + section geometry apply regardless of named styles — both
    // read attrs (default font / sectionProperties) that a styles-less document
    // still carries.
    this.#applyDefaultFontMetric();
    this.#applySectionGeometry();
  }

  /** Set a document-wide --docen-font-metric fallback on .docen-pages — the
   *  default font's `normal` ratio — so the page container's line-height
   *  (inherited by paragraphs without their own spacing) resolves to a real
   *  metric instead of the 1.2 fallback. Per-paragraph decorations override
   *  this for paragraphs that carry their own line-spacing. */
  #applyDefaultFontMetric(): void {
    const editor = this.#editor;
    const pages = this.shadowRoot?.querySelector<HTMLElement>(".docen-pages");
    if (!editor || !pages) return;
    const styles = (editor.state.doc.attrs?.styles ?? null) as StylesOptions | null;
    const { font } = effectiveRunProps(styles, null, {});
    const family = resolveFontName(font) ?? "serif";
    const ratio = fontNormalRatio({ family, bold: false, italic: false }).toFixed(4);
    pages.style.setProperty("--docen-font-metric", ratio);
  }

  /** Apply the document's section geometry — page size, orientation, and
   *  margins from `sectionProperties` (twips) — to the canvas, then
   *  re-paginate. Without this an imported document renders on the default
   *  A4 portrait + Normal margins instead of its real page setup, so the page
   *  count and breaks drift far from Word. twips → mm (1in = 1440tw = 25.4mm);
   *  canvas flips width/height for landscape via :host([orientation]). */
  #applySectionGeometry(): void {
    const editor = this.#editor;
    const canvas = this.shadowRoot?.querySelector("docen-document-area");
    if (!editor || !canvas) return;
    const sp = (editor.state.doc.attrs as Record<string, unknown> | undefined)
      ?.sectionProperties as
      | {
          page?: {
            size?: { width?: number; height?: number; orientation?: string };
            margin?: { top?: number; right?: number; bottom?: number; left?: number };
          };
          grid?: { linePitch?: number; type?: string } | null;
        }
      | undefined;
    const page = sp?.page;
    if (!page) return;
    // Page geometry in millimeters — the paper's natural unit (docx stores page
    // size/margins in twips; 1in = 1440tw = 25.4mm maps cleanly to mm). Font
    // sizes stay in pt (OOXML's unit); pt and mm are BOTH absolute CSS units
    // anchored to the same 96px/in reference pixel, so they render on one
    // consistent pixel grid — mm and pt do NOT need to be unified.
    if (page.size) {
      if (page.size.width) canvas.setAttribute("page-width", twipsToMm(page.size.width));
      if (page.size.height) canvas.setAttribute("page-height", twipsToMm(page.size.height));
    }
    // Render page-width × page-height directly. office-open's `orientation`
    // flag is unreliable (it can read "landscape" on portrait dimensions — this
    // very file is A4 portrait), so clear any prior landscape swap and let the
    // physical width/height decide the orientation.
    canvas.removeAttribute("orientation");
    if (page.margin) {
      const m = page.margin;
      const sides = [m.top, m.right, m.bottom, m.left];
      if (sides.every((s) => s != null))
        canvas.setAttribute("margin", sides.map(twipsToMm).join(" "));
    }
    // Document grid (w:docGrid) is applied PER PAGE: each page node renders its
    // section's linePitch as an inline line-height (page-node renderHTML), so
    // multi-section docs with different grids each render correctly. Nothing to
    // inject globally here — the per-page inline style cascades to the page's
    // paragraphs (line-height is inherited).
    this.#refreshGeometry();
  }

  /** The underlying Tiptap editor (for advanced, direct control). */
  getEditor(): Editor | undefined {
    return this.#editor;
  }

  /** Force a pagination re-measure now (bypasses the debounce). */
  repaginate(): void {
    if (this.#editor) pageStorageOf(this.#editor).repaginate();
  }

  // ── Task pane visibility (Office.addin.showAsTaskpane / hide equivalent) ──

  /** Show a task pane. No-op if already open. */
  showTaskpane(id: TaskPaneId): void {
    this.#setTaskpane(id, true);
  }

  /** Hide a task pane. No-op if already closed. */
  hideTaskpane(id: TaskPaneId): void {
    this.#setTaskpane(id, false);
  }

  /** Whether a task pane is currently open. */
  getTaskpaneState(id: TaskPaneId): boolean {
    return !!this.#paneEl(id)?.open;
  }

  #paneEl(id: TaskPaneId): (HTMLElement & { open: boolean }) | null {
    const pos = TASKPANE_POSITION[id];
    return this.shadowRoot?.querySelector(`docen-task-pane[position="${pos}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
  }

  /** Apply a visibility state and dispatch `docen:taskpane-visibility-change`
   *  when it flips. The detail carries `visibilityMode: "taskpane"|"hidden"` to
   *  mirror `Office.VisibilityMode`. Idempotent — no event when state is unchanged. */
  #setTaskpane(id: TaskPaneId, open: boolean): void {
    const pane = this.#paneEl(id);
    if (!pane || pane.open === open) return;
    pane.open = open;
    this.dispatchEvent(
      new CustomEvent("docen:taskpane-visibility-change", {
        bubbles: true,
        composed: true,
        detail: { id, visibilityMode: (open ? "taskpane" : "hidden") as VisibilityMode },
      }),
    );
  }

  // ── Zoom (method + event + getter; once `zoom` attr seeds #zoom) ──

  /** Apply a zoom level (percent, clamped 10–500). Idempotent; dispatches
   *  `docen:zoom-change` on a real change (mirrors `Office.Document.zoom.set`). */
  setZoom(pct: number): void {
    this.#setZoom(pct);
  }

  /** Current zoom level (percent). */
  getZoom(): number {
    return this.#zoom;
  }

  // ── Formatting marks (method + event; boolean `show-marks` attribute) ──

  /** Toggle editing/formatting marks on or off. Idempotent; dispatches
   *  `docen:marks-change`. The boolean `show-marks` attribute is the source of
   *  truth — CSS `:host([show-marks])` drives the page/section-break markers —
   *  so it's toggled directly. (`@attr({mode:"boolean"})` does not reflect
   *  property→attribute in fast-element 3.x, so a method beats a reactive attr
   *  here; see docen-ui-state-attribute-strategy.) */
  setShowMarks(on: boolean): void {
    if (this.hasAttribute("show-marks") === on) return;
    this.toggleAttribute("show-marks", on);
    this.#editor?.commands.toggleFormattingMarks();
    this.dispatchEvent(
      new CustomEvent("docen:marks-change", {
        bubbles: true,
        composed: true,
        detail: { showMarks: on },
      }),
    );
  }

  /** Whether editing/formatting marks are currently shown. */
  getShowMarks(): boolean {
    return this.hasAttribute("show-marks");
  }
}

export default DocenDocument;
