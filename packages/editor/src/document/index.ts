import {
  compileDocument,
  convertMillimetersToTwip,
  docxExtensions,
  effectiveRunProps,
  generateDOCX,
  generateMarkdown,
  normalizeDocument,
  parseDOCX,
  parseHTMLBody,
  parseMarkdown,
  sectionPageSizeDefaults,
  type JSONContent,
  type SectionPropertiesOptions,
  type StylesOptions,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import {
  projectDocumentOptions,
  type ProjectedFlowBox,
  type ProjectedPageBackground,
  type ProjectedPageFurniture,
  type ProjectedSection,
} from "@docen/docx/layout";
import {
  browserFontMetrics,
  layoutFlowSections,
  TextMeasurer,
  twipToPx,
  type FlowPage,
  type FlowPageInsets,
} from "@docen/layout";
import { attr, css, customElement, html } from "@microsoft/fast-element";
import type { Mark } from "@tiptap/pm/model";
import { EditorState, TextSelection, type Transaction } from "@tiptap/pm/state";
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
  notifyLocaleChange,
  observeLang,
  registerComponents,
  resolveTheme,
  t,
  type DocenAddin,
} from "../ui";
import { createDefaultAddin, textCounter } from "./addin";
import {
  mountEditBridge,
  type EditBridge,
  type StoryKind,
  type StorySlot,
} from "./canvas/edit-bridge";
// Side-effect: register the document-specific UI components moved out of the
// shared ui/ barrel — <docen-format-pane> (properties fallback) and
// <docen-outline> (navigation Headings tab).
import "./components/format-pane";
import "./components/outline";
import {
  CanvasStage,
  type CanvasStageSection,
  type LaidFurnitureSection,
  layFurnitureSections,
} from "./canvas/stage";
import type { OutlineItem } from "./components/outline";
import { WIRED_DISPATCH } from "./extensions/commands";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { renderRibbonFromSchema, ribbonActions, ribbonTabs } from "./ribbon";

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
function detectOpenFormat(file: File): "docx" | "markdown" {
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
const SAVE_FORMATS: Record<
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
  // Footnote prompts for the note text, references the caret and appends the
  // note body to documentExtras.footnotes.
  "insert-footnote",
  // Page Color writes the doc-level w:background (doc.attrs.background) from
  // the color-picker's palette value.
  "page-color",
  // Link prompts for an address and marks the selection (Word's Insert Link).
  "link",
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
  // #onChange (data-event)
  "open",
  "save-as",
  "print",
]);

const documentStyles = css`
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
  /* The canvas surface — the scroll container sits one level up (the
       document-area); this wrapper just anchors the edit bridge's textarea and
       caret overlays (position:relative). cursor:text is the editing surface's
       I-beam, like Word's page area. */
  .docen-canvas {
    position: relative;
    width: fit-content;
    margin: 0 auto;
    padding: 32px 0;
    cursor: text;
  }
  /* Grey the "Auto-save" label to match its disabled switch (skeleton
       feature), so the label + switch read as one unavailable control, like
       ribbon skeleton buttons. Lifts automatically once the switch loses
       disabled. */
  .autosave-label:has(+ fluent-switch[disabled]) {
    color: var(--docen-color-text-3, #8a8a8a);
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
      <div class="docen-canvas" part="page"></div>
    </docen-document-area>
    <docen-task-pane slot="task-pane-end" position="end" part="props-pane">
      <slot name="properties">
        <docen-format-pane></docen-format-pane>
      </slot>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="comments-pane" title="Comments">
      <docen-comments-pane></docen-comments-pane>
    </docen-task-pane>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <docen-options-dialog part="options"></docen-options-dialog>
  <docen-word-count-dialog part="word-count"></docen-word-count-dialog>
  <docen-symbol-dialog part="symbol"></docen-symbol-dialog>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <input type="file" id="file-input" accept=".docx,.md,.markdown" hidden />
  <input type="file" id="image-input" accept="image/*" hidden />
`;

/** Build a nested OutlineItem tree from the flat outline anchor list: each
 *  heading nests under the nearest preceding heading with a smaller level. */
function buildOutlineTree(
  anchors: readonly { id: string; textContent: string; originalLevel: number }[],
): OutlineItem[] {
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

/** MS Office margin presets (mm). `normal` matches the engine default
 *  (@office-open/docx sectionMarginDefaults: top/bottom 25.4mm, left/right
 *  31.75mm = MS Office zh-CN "Normal"). */
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
  const mergeGroup = <T extends object>(
    b: T | false | undefined,
    p: T | false | undefined,
  ): T | false | undefined =>
    p === undefined ? b : p === false || b === undefined || b === false ? p : { ...b, ...p };
  return {
    ...base,
    pageSize: mergeGroup(base?.pageSize, patch.pageSize),
    pageMargin: mergeGroup(base?.pageMargin, patch.pageMargin),
  };
}

/**
 * `<docen-document>` — a turnkey DOCX editor web component.
 *
 * Wires the Fluent UI host (title-bar + ribbon + document-area) to the canvas
 * route: a viewless Tiptap engine (the single source of truth for content and
 * commands) driving the layout pipeline (compile → project → layout → LeaferJS
 * paint) on every transaction. The title bar drives file I/O (open/save) and
 * language switching, ribbon commands route to the engine, and file I/O goes
 * through `parseDOCX`/`generateDOCX`. The title bar + ribbon re-render on
 * locale change.
 */

/**
 * Task pane identifiers, mirroring the Office `<TaskpaneId>` concept. The host
 * ships two built-in panes: `navigation` (start/left) and `properties` (end/right).
 */
export type TaskPaneId = "navigation" | "properties" | "comments";

/**
 * Visibility mode values, matching `Office.VisibilityMode` (`taskpane` | `hidden`).
 * Carried on {@link docen:taskpane-visibility-change} event details.
 */
export type VisibilityMode = "taskpane" | "hidden";

@customElement({ name: "docen-document", template: documentTemplate, styles: documentStyles })
class DocenDocument extends AddinHost<Editor> {
  // ── Reactive attributes (@attr) — no `reflect` (attribute → property stays
  //  one-way). addinsAttr (attribute "addins") dodges AddinHost.addinsChanged
  //  and the `addins` getter.
  @attr editable?: string;
  @attr filename?: string;
  @attr user?: string;
  @attr avatar?: string;
  @attr({ attribute: "section-properties" }) sectionProperties?: string;
  @attr styles?: string;
  @attr({ attribute: "addins" }) addinsAttr?: string;
  @attr theme?: string;

  #bridge?: EditBridge;
  #stage?: CanvasStage;
  #stageHost?: HTMLElement;
  #measurer = new TextMeasurer(browserFontMetrics);
  #pages: readonly FlowPage[] = [];
  /** Page index → section index (the caret's section and per-page geometry
   *  read through it). */
  #sectionOfPage: readonly number[] = [];
  /** The first section's flow box (page-width presets / TOC tab position;
   *  a multi-section refinement reads the caret's own section). */
  #flow?: ProjectedFlowBox;
  #fileInput?: HTMLInputElement;
  #imageInput?: HTMLInputElement;
  /** Latest TOC anchors, refreshed by the Outline extension; used to resolve
   *  an outline click back to a document position (pos). */
  #anchors: readonly { id: string; pos: number; textContent: string; originalLevel: number }[] = [];
  /** Cached doc nodeSize + Office-style word count so caret-move transactions
   *  don't re-walk the whole document (recomputed only when content changes). */
  #lastDocSize = -1;
  #lastWords = 0;
  /** Semantic fingerprint of the last outline tree — id/level/title only. `pos`
   *  shifts on every re-render but never changes what the pane shows, so it's
   *  excluded; the fingerprint is built from per-anchor arrays (not the
   *  serialized tree) so object key order can never cause a spurious mismatch. */
  #outlineSig = "";
  #unobserveLang?: () => void;
  /** Watches the host's `lang` attribute and forwards it to the internal
   *  <docen-workspace> + notifies locale observers. MutationObserver because
   *  @attr `lang` clashes with HTMLElement.lang (TS2416); manual
   *  observedAttributes would break FASTElement's @attr dispatch. */
  #langObserver?: MutationObserver;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured marks + the pointerup listener that applies them.
  #painterMarks: readonly Mark[] | null = null;
  #painterOff?: () => void;
  /** Current zoom level (percent) applied by the page stage's slot sizing. */
  #zoom = 100;
  /** Debounce timer for the nav-pane search result list — the list rebuilds only
   *  after the user pauses typing (the query dispatches immediately, so Enter /
   *  find-next stays in sync with the last keystroke). */
  #searchTimer?: ReturnType<typeof setTimeout>;
  /** The selection captured when a New Comment opened the compose box — the
   *  comment anchors there on Post (the box takes focus, moving the caret). */
  #pendingCommentRange?: { from: number; to: number };
  /** Cached unwrapped JSON (host.getJSON result). Invalidated on every user/doc
   *  change; recomputed lazily. Saves the editor.getJSON walk on every
   *  save/autosave/getJSON call. */
  #cachedJSON?: JSONContent;
  #jsonDirty = true;
  /** The header/footer story under edit (null = none). `#storyPage` is the
   *  anchor page the story edits in place on. */
  #storyKind: StoryKind | null = null;
  #storyPage = -1;

  /** The underlying Tiptap Editor (undefined before connect / after disconnect).
   *  Exposed so a host (the @docen/vue adapter, or any parent element) can drive
   *  commands programmatically — setContent / getJSON / chain / ... — without
   *  routing through the ribbon. */
  get editor(): Editor | undefined {
    return this.#bridge?.editor;
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

  // ── @attr change callbacks — every handler is guarded (bridge/shadowRoot
  //  check) so an early fire during FAST's attribute hydration is a no-op.
  editableChanged(): void {
    this.#bridge?.editor.setEditable(this.editable !== "false");
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
    if (ribbon) ribbon.setAttribute("mode", "always-shown");
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
  #renderOutline(
    anchors: readonly { id: string; pos: number; textContent: string; originalLevel: number }[],
  ): void {
    this.#anchors = anchors;
    const outline = this.shadowRoot?.querySelector("docen-outline");
    if (!outline) return;
    // Fingerprint only what the pane shows (id/level/title). `pos` moves on
    // every re-render but never changes the outline, so excluding it avoids
    // rebuilding — and flickering — the fluent tree each pass. Built from
    // per-anchor arrays rather than the serialized tree, so object key order
    // is irrelevant (no dependency on buildOutlineTree's literal field order,
    // unlike a plain JSON.stringify(tree) comparison).
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
    const bridge = this.#bridge;
    if (!id || !bridge) return;
    const anchor = this.#anchors.find((a) => a.id === id);
    if (!anchor) return;
    this.#setTextSelection(anchor.pos);
    bridge.scrollIntoView(anchor.pos);
  };

  /** navigation:search → set the active query; matches highlight live. */
  readonly #onSearch = (event: CustomEvent<{ query?: string }>): void => {
    const editor = this.editor;
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
    const editor = this.editor;
    if (!editor) return;
    (event.detail.direction === "prev" ? findPrev : findNext)(editor.state, editor.view.dispatch);
  };

  /** Stamp the Results slot with the live match list — each hit rendered with
   *  surrounding context and a data-from/to for click-to-jump (Word's Results
   *  pane lists every match with context, not just a count). */
  #updateSearchResults(): void {
    const editor = this.editor;
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
    const bridge = this.#bridge;
    if (!bridge) return;
    const item = (event.target as HTMLElement | null)?.closest(".result-item");
    if (!(item instanceof HTMLElement)) return;
    const from = Number(item.dataset.from);
    const to = Number(item.dataset.to);
    if (!Number.isFinite(from) || !Number.isFinite(to)) return;
    this.#setTextSelection(from, to);
    bridge.scrollIntoView(from);
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
    const editor = this.editor;
    if (!editor) return;
    const { action, find, replace, caseSensitive, wholeWord } = event.detail ?? {};
    const query = new SearchQuery({ search: find, replace, caseSensitive, wholeWord });
    editor.view.dispatch(setSearchState(editor.state.tr, query));
    if (action === "find-next") findNext(editor.state, editor.view.dispatch);
    else if (action === "replace-next") replaceNext(editor.state, editor.view.dispatch);
    else if (action === "replace-all") replaceAll(editor.state, editor.view.dispatch);
  };

  /** Set a text selection (or a range) on the viewless editor. Same runtime
   *  PM instance — the cast bridges the dual d.ts identity between this
   *  package's @tiptap/pm and the engine's. */
  #setTextSelection(from: number, to?: number): void {
    const editor = this.editor;
    if (!editor) return;
    const sel = TextSelection.create(editor.state.doc, from, to);
    editor.view.dispatch(editor.state.tr.setSelection(sel as never));
    this.#bridge?.focus();
  }

  /** Paste from the system clipboard. Prefers text/html — styled paste through
   *  the schema's parse rules — and falls back to plain text; `textOnly` (the
   *  menu's Keep Text Only) skips the HTML leg. navigator.clipboard is the
   *  reliable path; execCommand("paste") is the fallback (often blocked). */
  async #paste(textOnly = false): Promise<void> {
    const editor = this.editor;
    if (!editor) return;
    this.#bridge?.focus();
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const type =
          !textOnly && item.types.includes("text/html")
            ? "text/html"
            : item.types.includes("text/plain")
              ? "text/plain"
              : null;
        if (!type) continue;
        const text = await (await item.getType(type)).text();
        if (!text) continue;
        if (type === "text/html") {
          const body = new DOMParser().parseFromString(text, "text/html").body;
          const content = parseHTMLBody(body, editor.state.schema).content ?? [];
          if (content.length) {
            editor.commands.insertContent(content);
            return;
          }
        } else {
          editor.commands.insertContent(text);
          return;
        }
      }
    } catch {
      // read() may be denied (permission policy) — fall through to readText.
    }
    try {
      const text = await navigator.clipboard.readText();
      if (text) editor.commands.insertContent(text);
    } catch {
      /* clipboard unavailable — nothing to paste */
    }
  }

  /** Editing → Select menu. "all" uses the official selectAll() command;
   *  "objects"/"similar" are placeholders. */
  #select(value?: string): void {
    const editor = this.editor;
    if (!editor) return;
    if ((value ?? "all") !== "all") return;
    this.#bridge?.focus();
    editor.commands.selectAll();
  }

  /** Editing → Find drop-down → Go To: prompt for a page number and move the
   *  caret to that page, scrolling it into view. */
  #goToPage(): void {
    const bridge = this.#bridge;
    if (!bridge) return;
    const input = window.prompt(t("ribbon.opt.go-to-prompt", this));
    if (input == null) return;
    const page = parseInt(input, 10);
    if (!Number.isFinite(page) || page < 1 || page > this.#pages.length) return;
    const pos = bridge.firstPosOfPage(page - 1);
    if (pos == null) return;
    this.#setTextSelection(pos);
    bridge.scrollIntoView(pos);
  }

  /** Format Painter: on first click, capture the current selection's marks and
   *  arm a one-shot pointerup listener; the next non-empty selection receives
   *  those marks and disarms the painter. A second click cancels. */
  #toggleFormatPainter(): void {
    if (this.#painterMarks) {
      this.#stopFormatPainter();
      return;
    }
    const editor = this.editor;
    if (!editor || editor.state.selection.empty) return;
    // Probe one character into the selection: $from sits on the boundary,
    // and ResolvedPos.marks() reads the character BEFORE the position — the
    // first selected character's marks (e.g. bold stamped on [from,to))
    // would be lost.
    this.#painterMarks = editor.state.doc.resolve(editor.state.selection.from + 1).marks();
    this.toggleAttribute("format-painter", true);
    const onUp = (): void => {
      const ed = this.editor;
      if (!ed) return;
      const { from, to, empty } = ed.state.selection;
      if (!empty && this.#painterMarks) {
        const tr = ed.state.tr;
        for (const mark of this.#painterMarks) tr.addMark(from, to, mark);
        ed.view.dispatch(tr);
      }
      this.#stopFormatPainter();
    };
    this.addEventListener("pointerup", onUp, { once: true });
    this.#painterOff = () => this.removeEventListener("pointerup", onUp);
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
   *  moves, marks change). */
  #setupFontSync(): void {
    const editor = this.editor;
    if (!editor) return;
    const sync = (): void => {
      this.#syncFontControls();
      this.#syncStyleControl();
      this.#syncStoryMenus();
      this.#updateStatus();
    };
    editor.on("transaction", sync);
    sync();
    this.#fontSyncCleanup = (): void => {
      editor.off("transaction", sync);
    };
  }

  #syncFontControls(): void {
    const editor = this.editor;
    if (!editor) return;
    // Resolve font + size in one pass through the style inheritance chain
    // (direct run props → paragraph style → basedOn → document defaults).
    const { font, size } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      editor.getAttributes("textStyle"),
    );
    const fontDisplay = font ?? "";
    const sizeDisplay = size != null ? String(size) : "";
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

  /** The loaded document's styles model (doc.attrs.styles), or null. */
  #docStyles(editor: Editor): StylesOptions | null {
    return (editor.state.doc.attrs?.styles as StylesOptions | undefined) ?? null;
  }

  /** The paragraph-style id at the caret (the HeadingLevel literal carried on
   *  `heading` for heading paragraphs, the pStyle id on `style` otherwise). */
  #currentStyleId(editor: Editor): string | null {
    const attrs = editor.getAttributes("paragraph") as {
      heading?: unknown;
      style?: unknown;
    };
    if (typeof attrs.heading === "string" && attrs.heading) return attrs.heading;
    if (typeof attrs.style === "string" && attrs.style) return attrs.style;
    return null;
  }

  /** Mirror the paragraph style at the caret into the Styles gallery combobox —
   *  its value is the current paragraph's style id (the HeadingLevel literal
   *  carried on `heading` for heading paragraphs, the pStyle id on `style`
   *  otherwise, or "Normal" when the paragraph carries none). The combobox
   *  matches the value against its gallery items to show the style's name. */
  #syncStyleControl(): void {
    const editor = this.editor;
    if (!editor) return;
    const value = this.#currentStyleId(editor) || "Normal";
    const cb = this.shadowRoot?.querySelector<HTMLElement>('docen-ribbon-combobox[event="style"]');
    if (cb && cb.getAttribute("value") !== value) cb.setAttribute("value", value);
  }

  async connectedCallback(): Promise<void> {
    super.connectedCallback();
    // Forward this host's `lang` attribute to the internal <docen-workspace>
    // so resolveLang (scoped to docen-workspace) honors <docen-document lang>,
    // not just <html lang>. See #syncLang for why MutationObserver, not @attr.
    this.#langObserver = new MutationObserver(() => this.#syncLang());
    this.#langObserver.observe(this, { attributes: true, attributeFilter: ["lang"] });
    this.#syncLang();
    await registerComponents();
    applyTheme(resolveTheme(this.getAttribute("theme")));

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

    this.#stageHost = this.shadowRoot!.querySelector<HTMLElement>(".docen-canvas") ?? undefined;
    if (!this.#stageHost) return;

    // Fonts must be loaded before the pipeline measures, else the layout
    // drifts from the browser's actual font metrics.
    await document.fonts?.ready;

    const contentAttr = this.getAttribute("content");
    // Declarative section-properties / styles (JSON) seed doc-level attrs so a
    // host can bootstrap page setup + named styles without openDOCX/setJSON.
    const initAttrs = this.#readInitAttrs();
    // The content attribute accepts Tiptap JSON only; a malformed string
    // mounts an empty document rather than throwing mid-connection.
    let baseDoc = {} as JSONContent;
    if (contentAttr) {
      try {
        baseDoc = JSON.parse(contentAttr) as JSONContent;
      } catch {
        console.warn("[docen-document] content attribute is not valid JSON — ignored");
      }
    }
    const seeded =
      Object.keys(initAttrs).length > 0
        ? { ...baseDoc, attrs: { ...baseDoc.attrs, ...initAttrs } }
        : baseDoc;
    // Fill office-open's document-level defaults (the built-in style library +
    // page geometry + docGrid linePitch) so a freshly mounted document matches
    // what export produces. Host-declared initAttrs win — normalizeDocument
    // shallow-merges user attrs over the defaults.
    const initialDoc =
      (seeded.attrs as { sectionProperties?: unknown } | undefined)?.sectionProperties == null
        ? normalizeDocument(seeded)
        : seeded;
    // The default document add-in contributes the engine extensions + every
    // wired ribbon command. Registered before the editor mounts so its
    // extensions seed the schema. Ribbon events route straight to the engine
    // (editor.commands.<event>), not addin.commands.
    const defaultAddin = createDefaultAddin({
      onOutlineUpdate: (anchors) => this.#renderOutline(anchors),
    });
    this.addAddin(defaultAddin);
    // Declarative external add-ins (JSON `addins` attribute) register after the
    // default so their ribbon tabs append to the built-ins via mergeRibbonSchema.
    this.#applyAddinsAttr();

    this.#bridge = mountEditBridge({
      host: this.#stageHost,
      content: initialDoc,
      onDoc: (json) => this.#renderDoc(json),
      pageHost: (page) => this.#stage?.slotAt(page)?.parentElement ?? null,
      extensions: [...docxExtensions, ...(defaultAddin.extensions ?? [])],
      scale: () => this.#stage?.scale() ?? 1,
      // Header/footer edit stories — the bridge routes band double-clicks
      // here; the host owns the slots' persistence (attrs on the doc node or
      // the section's last paragraph).
      story: {
        geometry: (kind, page) => {
          const stage = this.#stage;
          const band = stage?.furnitureBand(kind, page);
          if (!stage || !band) return null;
          return {
            stack: stage.furnitureStack(kind, page),
            band,
            slot: stage.slotOfPage(page),
          };
        },
        read: (kind, slot, page) => this.#readStorySource(kind, slot, page),
        entered: (kind, slot, page) => {
          this.#storyKind = kind;
          this.#storyPage = page;
          this.#stage?.setStoryEdit({
            kind,
            label: t(kind === "header" ? "story.header" : "story.footer", this),
          });
        },
        onDoc: (kind, slot, json) => this.#renderStoryFurniture(kind, slot, json),
        exit: ({ kind, slot, json, dirty }) => this.#exitStory(kind, slot, json, dirty),
      },
      // Drawing selection — the stage's painted-box hit table resolves the
      // click; the caret map pairs the hit's host paragraph to the PM
      // position, and the index-th drawing node inside it becomes the
      // NodeSelection (projectDrawings collects drawings in run order, the
      // same order the paragraph's content carries the nodes).
      drawingAt: (page, lx, ly) => this.#stage?.drawingAt(page, lx, ly) ?? null,
      drawingSelection: (hit) => this.#drawingNodePos(hit.para, hit.index, hit.kind),
      drawingBoxOf: (para, index, kind) => this.#stage?.drawingBoxOf(para, index, kind) ?? null,
    });
    if (this.getAttribute("editable") === "false") this.#bridge.editor.setEditable(false);
    // First paint + caret map feed (transactions re-render via the bridge's
    // raf-merged onDoc from here on).
    this.#renderDoc(initialDoc);

    // Mirror the caret's font/size into the ribbon comboboxes (Word behavior).
    this.#setupFontSync();

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
    // Comments pane → create/cancel (compose box), select (scroll to range),
    // update/delete (inline card actions).
    this.addEventListener("comment:create", this.#onCommentCreate as EventListener);
    this.addEventListener("comment:cancel", this.#onCommentCancel as EventListener);
    this.addEventListener("comment:select", this.#onCommentSelect as EventListener);
    this.addEventListener("comment:update", this.#onCommentUpdate as EventListener);
    this.addEventListener("comment:delete", this.#onCommentDelete as EventListener);
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
    // Symbol dialog — insert the picked character at the caret.
    this.shadowRoot!.querySelector("docen-symbol-dialog")?.addEventListener(
      "symbol:insert",
      this.#onSymbolInsert as EventListener,
    );

    // Re-render header + ribbon when the page locale (<html lang>) changes.
    this.#unobserveLang = observeLang(() => this.#renderChrome());

    // Ribbon Display Options → drive browser fullscreen + status-bar hide.
    // auto-hide = Full Screen (Office); any other mode exits it.
    const ribbon = this.shadowRoot!.querySelector("docen-ribbon");
    ribbon?.addEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    // Emit docen:change on every content change (autosave driver) and docen:ready
    // once the editor is live — both bubble out so a host can react.
    this.editor?.on("transaction", this.#onTransaction);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
    this.dispatchEvent(new CustomEvent("docen:ready", { bubbles: true, composed: true }));
  }

  #slotsKeyOf(kind: StoryKind): "sectionHeaders" | "sectionFooters" {
    return kind === "header" ? "sectionHeaders" : "sectionFooters";
  }

  /** The story's source JSON — the section owning `page` holds the slots
   *  (Word: the band double-clicked edits that page's section, regardless of
   *  where the caret sits); an absent slot falls back to the default slot's
   *  (what the page displays until the edit breaks the tie). */
  #readStorySource(kind: StoryKind, slot: StorySlot, page: number): JSONContent[] {
    const editor = this.editor;
    if (!editor) return [];
    const pos = this.#sectPrPosOfSection(this.#sectionOfPage[page] ?? 0);
    const attrs =
      pos >= 0
        ? (editor.state.doc.nodeAt(pos)?.attrs as Record<string, unknown> | undefined)
        : (editor.state.doc.attrs as Record<string, unknown> | undefined);
    const group = attrs?.[this.#slotsKeyOf(kind)] as
      | { default?: JSONContent[]; first?: JSONContent[]; even?: JSONContent[] }
      | undefined;
    return group?.[slot] ?? group?.default ?? [];
  }

  /** A story keystroke's render path: patch the slot into a copy of the doc
   *  JSON and re-run the full pipeline — the body re-flows because the
   *  header/footer it edits pushes on it (Word: typing in a header moves the
   *  body live). The fresh furniture stack goes back to the story's map. */
  #renderStoryFurniture(kind: StoryKind, slot: StorySlot, json: JSONContent[]): void {
    const bridge = this.#bridge;
    const stage = this.#stage;
    if (!bridge || !stage || this.#storyKind !== kind) return;
    const raw = bridge.editor.getJSON();
    // getJSON()'s attrs object IS the live PM attrs (Node.toJSON carries it by
    // reference) — patch a shallow copy or the editor state would mutate
    // without a transaction (no render, no undo, no docen:change).
    const attrs = { ...(raw.attrs as Record<string, unknown>) };
    const key = this.#slotsKeyOf(kind);
    attrs[key] = { ...(attrs[key] as object | undefined), [slot]: json };
    const run = this.#projectAndLayout({ ...raw, attrs } as JSONContent);
    this.#pages = run.pages;
    this.#sectionOfPage = run.sectionOfPage;
    this.#flow = run.sections[0]?.flow;
    stage.sync(run.pages, run.sections, run.sectionOfPage, run.background);
    bridge.updatePages(run.pages, this.#pageOriginOf(run.sections, run.sectionOfPage));
    const band = stage.furnitureBand(kind, this.#storyPage);
    bridge.updateStoryMap(
      band ? stage.furnitureStack(kind, this.#storyPage) : null,
      band ?? { top: 0, bottom: 0, paintY: 0 },
    );
  }

  /** The doc position of the paragraph closing the given section (0-based —
   *  the Nth sectionProperties paragraph in document order), or -1 when that
   *  section closes at the body end (its sectPr lives on the doc node). */
  #sectPrPosOfSection(sectionIndex: number): number {
    const editor = this.editor;
    if (!editor) return -1;
    let seen = -1;
    let target = -1;
    editor.state.doc.descendants((node, pos) => {
      if (target >= 0) return false;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        seen++;
        if (seen === sectionIndex) {
          target = pos;
          return false;
        }
      }
      return true;
    });
    return target;
  }

  /** Write a finished story's JSON back. The story edits the section its
   *  anchor page belongs to — the caret is no address here (exiting by
   *  clicking another page's body moves it). An earlier section's slots live
   *  on its closing sectPr paragraph and go through a plain setNodeMarkup
   *  transaction — one undo step. The final section closes at the body end
   *  and its slots live on the doc node, which no step can address
   *  (nodeAt(0) is the first child) — they land through #loadDoc's state
   *  rebuild, the same path setJSON takes (history resets with it, like any
   *  document load). */
  #persistStory(kind: StoryKind, slot: StorySlot, json: JSONContent[], anchorPage: number): void {
    const bridge = this.#bridge;
    if (!bridge) return;
    const key = this.#slotsKeyOf(kind);
    const slots = (attrs: Record<string, unknown>): Record<string, unknown> => ({
      ...(attrs[key] as object | undefined),
      [slot]: json,
    });
    const sectionIndex = this.#sectionOfPage[anchorPage] ?? 0;
    const target = this.#sectPrPosOfSection(sectionIndex);
    if (target < 0) {
      const raw = bridge.editor.getJSON();
      this.#loadDoc({
        ...raw,
        attrs: {
          ...(raw.attrs as Record<string, unknown>),
          [key]: slots(raw.attrs as Record<string, unknown>),
        },
      } as JSONContent);
      return;
    }
    bridge.editor.commands.command(({ state: s, dispatch }) => {
      const node = s.doc.nodeAt(target)!;
      dispatch?.(
        s.tr.setNodeMarkup(target, undefined, { ...node.attrs, [key]: slots(node.attrs) }),
      );
      return true;
    });
  }

  #exitStory(kind: StoryKind, slot: StorySlot, json: JSONContent[], dirty: boolean): void {
    this.#stage?.setStoryEdit(null);
    this.#storyKind = null;
    if (dirty) this.#persistStory(kind, slot, json, this.#storyPage);
    this.#storyPage = -1;
  }

  /** Write a slots group through a transaction: the group lives on the
   *  current section's sectPr paragraph when there is one, else on the doc
   *  node (the #loadDoc state-rebuild path — the doc node is not step
   *  addressable; see #persistStory). */
  #writeSlots(key: "sectionHeaders" | "sectionFooters", group: Record<string, unknown>): void {
    const bridge = this.#bridge;
    const editor = this.editor;
    if (!bridge || !editor) return;
    const { doc, tr } = editor.state;
    const targetPos = this.#sectionSectPrPos();
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        tr.setNodeMarkup(targetPos, undefined, { ...node.attrs, [key]: group });
        editor.view.dispatch(tr);
        return;
      }
    }
    const raw = bridge.editor.getJSON();
    this.#loadDoc({
      ...raw,
      attrs: { ...(raw.attrs as Record<string, unknown>), [key]: group },
    } as JSONContent);
  }

  /** Remove Header / Remove Footer — drop the story's whole slots group from
   *  the current section (Word removes the content; the slot stops
   *  rendering on every page). */
  #removeStory(kind: StoryKind): void {
    this.#writeSlots(this.#slotsKeyOf(kind), {});
  }

  /** Is this inline passthrough atom a PAGE (not NUMPAGES/PAGEREF) field? */
  static isPageField(child: JSONContent): boolean {
    if (child.type !== "inlinePassthrough") return false;
    try {
      const data = JSON.parse(String((child.attrs as { data?: string } | undefined)?.data)) as {
        simpleField?: { instruction?: string };
      };
      const instr = data.simpleField?.instruction?.trim().toUpperCase() ?? "";
      return instr.startsWith("PAGE") && !instr.startsWith("PAGES") && !instr.startsWith("PAGEREF");
    } catch {
      return false;
    }
  }

  /** The slots group as #writeSlots addresses it (same container semantics:
   *  the current section's sectPr paragraph, else the doc node). */
  #readSlotsGroup(key: "sectionHeaders" | "sectionFooters"): Record<string, unknown> {
    const editor = this.editor;
    if (!editor) return {};
    const targetPos = this.#sectionSectPrPos();
    const group = (attrs: Record<string, unknown> | undefined): Record<string, unknown> =>
      (attrs?.[key] as Record<string, unknown> | undefined) ?? {};
    if (targetPos != null) {
      const node = editor.state.doc.nodeAt(targetPos);
      if (node) return group(node.attrs as Record<string, unknown>);
    }
    return group(editor.state.doc.attrs as Record<string, unknown>);
  }

  /** Remove Page Numbers — strip the PAGE field atoms from every slot of
   *  both stories (Word deletes the fields, leaving their paragraphs). */
  #removePageNumbers(): void {
    const strip = (blocks: unknown): unknown => {
      const json = blocks as JSONContent[] | undefined;
      if (!Array.isArray(json)) return blocks;
      return json.map((block) =>
        block.type === "paragraph"
          ? {
              ...block,
              content: (block.content ?? []).filter((c) => !DocenDocument.isPageField(c)),
            }
          : block,
      );
    };
    for (const key of ["sectionHeaders", "sectionFooters"] as const) {
      const group = this.#readSlotsGroup(key);
      const next: Record<string, unknown> = {};
      for (const slot of ["default", "first", "even"] as const) {
        if (group[slot] !== undefined) next[slot] = strip(group[slot]);
      }
      this.#writeSlots(key, next);
    }
  }

  /** Word's furniture overflow rule: a header taller than the top margin
   *  pushes the body down, a taller footer pushes it up — each page by its
   *  own slot's LAID stack (the first page by the first slot when titlePage
   *  asks for one, even pages by the even slot). Slots without their own
   *  content fall back to the default stack (OOXML reference semantics),
   *  matching the stage's paint fallback — the heights are the same layout
   *  pass the painter's bands come from. */
  #pageInsets(
    flow: ProjectedFlowBox,
    furniture: ProjectedPageFurniture | undefined,
    laid: LaidFurnitureSection | undefined,
  ): FlowPageInsets | undefined {
    if (!furniture) return undefined;
    const topMargin = flow.contentTopPx;
    const bottomMargin = flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx;
    const headerDistance = furniture.headerDistancePx ?? 48;
    const footerDistance = furniture.footerDistancePx ?? 48;
    const height = (kind: "header" | "footer", slot: 0 | 1 | 2): number | undefined =>
      laid?.[kind][slot]?.heightPx;
    const inset = (headerPx: number | undefined, footerPx: number | undefined) => {
      const top = Math.max(0, headerDistance + (headerPx ?? 0) - topMargin);
      const bottom = Math.max(0, footerDistance + (footerPx ?? 0) - bottomMargin);
      return top > 0 || bottom > 0
        ? { topPx: Math.round(top), bottomPx: Math.round(bottom) }
        : undefined;
    };
    const def = inset(height("header", 0), height("footer", 0));
    if (!def) return undefined;
    const out: FlowPageInsets = { default: def };
    if (furniture.titlePage) {
      out.first =
        inset(
          height("header", 1) ?? height("header", 0),
          height("footer", 1) ?? height("footer", 0),
        ) ?? undefined;
    }
    if (furniture.evenAndOddHeaders) {
      out.even =
        inset(
          height("header", 2) ?? height("header", 0),
          height("footer", 2) ?? height("footer", 0),
        ) ?? undefined;
    }
    return out;
  }

  /** The PM node position of a drawing hit's target — the host paragraph's
   *  inner position via the caret map, then the index-th node of the hit's
   *  kind: "drawing" counts floating pictures + wps shapes (projectDrawings'
   *  run order = the paragraph's content order), "inline" counts the
   *  paragraph's non-floating images (the line items' picture order). */
  #drawingNodePos(para: unknown, index: number, kind: "drawing" | "inline"): number | null {
    const bridge = this.#bridge;
    const doc = bridge?.editor.state.doc;
    const innerPos = bridge?.posOfPara(para) ?? null;
    if (innerPos == null || !doc) return null;
    const host = doc.nodeAt(innerPos - 1);
    if (!host) return null;
    let seen = 0;
    let hit = -1;
    host.forEach((child, offset) => {
      const target =
        kind === "drawing"
          ? child.type.name === "wpsShape" ||
            (child.type.name === "image" && child.attrs.floating != null)
          : child.type.name === "image" && child.attrs.floating == null;
      if (target && hit < 0 && seen++ === index) hit = innerPos + offset;
    });
    return hit >= 0 ? hit : null;
  }

  /** The canvas pipeline's projection + layout half, shared by the full
   *  render and the story's live re-render: compile → project (one section
   *  per document section) → lay each section's furniture ONCE (the insets
   *  and the painter's bands share the pass) → paginate continuously across
   *  sections (each section starts a fresh page; the page→section map drives
   *  per-page geometry everywhere). */
  #projectAndLayout(doc: JSONContent): {
    pages: FlowPage[];
    sectionOfPage: number[];
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
  } {
    const { sections, background } = projectDocumentOptions(compileDocument(doc));
    const stageSections: (ProjectedSection & CanvasStageSection)[] = sections.map((section) => ({
      ...section,
    }));
    const laidFurniture = layFurnitureSections(stageSections, browserFontMetrics);
    stageSections.forEach((section, i) => {
      section.furnitureLaid = laidFurniture[i];
    });
    const flowSections = stageSections.map((section) => {
      const pageInsets = this.#pageInsets(section.flow, section.furniture, section.furnitureLaid);
      return {
        blocks: section.blocks,
        opts: pageInsets ? { ...section.flow, pageInsets } : section.flow,
      };
    });
    const { pages, sectionOfPage } = layoutFlowSections(flowSections, this.#measurer);
    return { pages, sectionOfPage, sections: stageSections, background };
  }

  /** The page→section origin resolver the bridge's caret maps need (each
   *  page's own section's content-box origin). */
  #pageOriginOf(
    sections: readonly (ProjectedSection & CanvasStageSection)[],
    sectionOfPage: readonly number[],
  ): (page: number) => { contentLeftPx: number; contentTopPx: number } {
    return (page) =>
      sections[sectionOfPage[page] ?? 0]?.flow ?? { contentLeftPx: 0, contentTopPx: 0 };
  }

  /** The canvas pipeline — the single render entry the bridge's transactions
   *  and the loaders share: compile → project → layout → paint, then re-arm
   *  the caret map against the fresh geometry. */
  #renderDoc(doc: JSONContent): void {
    if (!this.#stageHost) return;
    const run = this.#projectAndLayout(doc);
    this.#pages = run.pages;
    this.#sectionOfPage = run.sectionOfPage;
    this.#flow = run.sections[0]?.flow;
    this.#stage ??= new CanvasStage(this.#stageHost, {
      metrics: browserFontMetrics,
      sections: run.sections,
      sectionOfPage: run.sectionOfPage,
      background: run.background,
    });
    // A `zoom` attribute parsed before the stage existed only recorded the
    // level here — push it in before the first sync sizes the slots.
    if (this.#stage.zoom !== this.#zoom) this.#stage.setZoom(this.#zoom);
    this.#stage.sync(run.pages, run.sections, run.sectionOfPage, run.background);
    this.#bridge?.updatePages(run.pages, this.#pageOriginOf(run.sections, run.sectionOfPage));
    this.#updateStatus();
    this.#syncCommentsPane();
  }

  disconnectedCallback(): void {
    this.#langObserver?.disconnect();
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
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
      ?.querySelector("docen-symbol-dialog")
      ?.removeEventListener("symbol:insert", this.#onSymbolInsert as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.removeEventListener("zoom:change", this.#onZoomChange as EventListener);
    this.editor?.off("transaction", this.#onTransaction);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.shadowRoot
      ?.querySelector("docen-ribbon")
      ?.removeEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    this.#fontSyncCleanup?.();
    this.#fontSyncCleanup = undefined;
    this.#stopFormatPainter();
    clearTimeout(this.#searchTimer);
    this.#bridge?.destroy();
    this.#bridge = undefined;
    this.#stage?.destroy();
    this.#stage = undefined;
    super.disconnectedCallback();
  }

  #renderHeader(): string {
    const user = this.getAttribute("user") ?? "";
    const avatar = this.getAttribute("avatar") ?? "";
    const filename = this.getAttribute("filename") ?? t("header.doc-name", this);
    const initial = user.trim().charAt(0).toUpperCase();
    const avatarMarkup = avatar
      ? `<img class="avatar avatar-img" src="${escapeHtml(avatar)}" alt="" />`
      : initial
        ? `<span class="avatar">${initial}</span>`
        : "";
    const autosave = t("header.autosave", this);
    return `
          <div slot="start" style="display:flex;align-items:center;gap:4px">
            <span style="font-weight:600;font-size:13px;padding-inline:6px">${t("header.brand", this)}</span>
            <span class="autosave-label">${autosave}</span>
            <!-- auto-save is skeleton-only — disabled (greyed, non-interactive)
                 until the feature is wired; the autosave case in onChange is a
                 no-op. Remove disabled (and re-add checked) when it lands. -->
            <fluent-switch data-event="autosave" disabled aria-label="${autosave}"></fluent-switch>
            <docen-ribbon-button icon="save" label="${t("header.save", this)}" event="save" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="undo" label="${t("header.undo", this)}" event="undo" icon-only></docen-ribbon-button>
            <docen-ribbon-button icon="redo" label="${t("header.redo", this)}" event="redo" icon-only></docen-ribbon-button>
            <fluent-menu>
              <fluent-menu-button slot="trigger" appearance="subtle">${escapeHtml(filename)}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new">${t("header.new", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="open">${t("header.open", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="save-as">${t("header.save-as", this)}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-markdown">${t("header.save-as-markdown", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="print">${t("header.print", this)}</fluent-menu-item>
                <fluent-menu-item data-event="options">${t("header.options", this)}</fluent-menu-item>
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
    const styles = this.editor?.state.doc.attrs?.styles ?? null;
    titleBar.innerHTML = this.#renderHeader();
    // Built-in tabs (Home/Insert/… with the live style gallery) come from
    // ribbonTabs; external add-ins layer their own tabs on top via
    // mergeRibbonSchema. The default add-in contributes no ribbon, so without
    // extra add-ins this is just the built-in set.
    const tabs = [...ribbonTabs(styles), ...mergeRibbonSchema(this.addins)];
    const ribbonEl = root.querySelector("docen-ribbon")!;
    // Pass the workspace as the i18n scope so labels resolve against
    // `<docen-workspace lang>` (forwarded from `<docen-document lang>`)
    // rather than `<html lang>`. `closest()` can't reach the workspace from
    // inside this fragment (shadow boundary + not yet inserted), so the
    // workspace element must be handed in explicitly.
    ribbonEl.replaceChildren(
      renderRibbonFromSchema(
        tabs,
        ribbonActions(),
        root.querySelector("docen-workspace") ?? document.documentElement,
      ),
    );
    // Feed the full ribbon schema (built-in tabs + addin contributions) to the
    // command search so it can flatten and index every command. Re-runs on
    // lang/addin change since #renderChrome is the single chrome re-stamp.
    const searchEl = root.querySelector("docen-command-search") as
      | (HTMLElement & { setTabs(tabs: readonly unknown[], scope: Element | null): void })
      | null;
    // Pass the workspace as the i18n scope so command labels resolve against
    // `<docen-workspace lang>` (forwarded from `<docen-document lang>`) — the
    // same scope the ribbon uses just above.
    searchEl?.setTabs(tabs, root.querySelector("docen-workspace"));
    this.#applyRibbonGreying();
    this.#syncEditModeMenu();
    this.#syncStoryMenus();
    root
      .querySelector('docen-task-pane[part="comments-pane"]')
      ?.setAttribute("title", t("comments.title", this));
    this.#renderPanes();
  }

  /** Addin registry changed (add-in registered/removed) — re-stamp the ribbon
   *  so an external add-in's ribbon contribution appears. The default add-in
   *  contributes no ribbon, so this is a no-op for it; only extra add-ins add
   *  tabs. */
  protected addinsChanged(): void {
    this.#renderChrome();
  }

  /** Stamp pane titles + status text for the active locale (re-run on lang change). */
  #renderPanes(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const navPane = root.querySelector('docen-task-pane[position="start"]');
    if (navPane) navPane.setAttribute("title", t("pane.navigation", this));
    const propsPane = root.querySelector('docen-task-pane[position="end"]');
    if (propsPane) propsPane.setAttribute("title", t("pane.properties", this));
    // Status bar is dynamic (page count / caret page / zoom) — re-stamp it so a
    // locale change re-localizes the text too.
    this.#updateStatus();
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

  /** Apply the `theme` attribute: switch the Fluent theme
   *  (light/dark/high-contrast/teams-*). */
  #applyThemeAttr(value: string): void {
    applyTheme(resolveTheme(value));
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
    const editable = this.editor?.isEditable ?? true;
    menu.setAttribute("label", t(editable ? "ribbon.opt.editing" : "ribbon.opt.viewing", this));
    menu.setAttribute(
      "items",
      JSON.stringify([
        {
          text: t("ribbon.opt.editing", this),
          event: "edit-mode",
          value: "edit",
          checked: editable,
        },
        {
          text: t("ribbon.opt.viewing", this),
          event: "edit-mode",
          value: "view",
          checked: !editable,
        },
      ]),
    );
  }

  /** Re-stamp the Header/Footer split drop-downs with live checked flags —
   *  the slot-visibility items read sectionProperties (titlePage /
   *  evenAndOddHeaders), which the static ribbon schema can't carry. Runs on
   *  every chrome re-stamp and every transaction (a flag toggle flips its
   *  check on the next pass). */
  #syncStoryMenus(): void {
    const sp = (
      this.editor?.state.doc.attrs as
        | { sectionProperties?: { titlePage?: boolean; evenAndOddHeaders?: boolean } }
        | undefined
    )?.sectionProperties;
    const stamp = (kind: "header" | "footer"): void => {
      const el = this.shadowRoot?.querySelector(`docen-ribbon-split-button[event="${kind}"]`);
      if (!el) return;
      el.setAttribute(
        "items",
        JSON.stringify([
          {
            text: t(kind === "header" ? "ribbon.opt.edit-header" : "ribbon.opt.edit-footer", this),
            value: "edit",
          },
          {
            text: t(
              kind === "header" ? "ribbon.opt.remove-header" : "ribbon.opt.remove-footer",
              this,
            ),
            value: kind === "header" ? "remove-header" : "remove-footer",
          },
          {
            text: t("ribbon.opt.different-first", this),
            value: "title-page",
            checked: !!sp?.titlePage,
          },
          {
            text: t("ribbon.opt.odd-even", this),
            value: "odd-even",
            checked: !!sp?.evenAndOddHeaders,
          },
        ]),
      );
    };
    stamp("header");
    stamp("footer");
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
    detail?: { format?: "docx" | "markdown" },
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
  readonly #onTransaction = (props: { transaction: Transaction }): void => {
    if (props.transaction.docChanged) {
      this.#jsonDirty = true;
      this.dispatchEvent(
        new CustomEvent("docen:change", { bubbles: true, composed: true, detail: { dirty: true } }),
      );
    }
  };

  /** Toggle a task pane open/closed (ribbon View → toggle-navigation). */
  #togglePane(id: TaskPaneId): void {
    this.#setTaskpane(id, !this.getTaskpaneState(id));
  }

  /** Apply a paper-size preset (a4/letter/…) — writes the size into the
   *  document-model sectionProperties (Word stores page setup in the sectPr)
   *  so layout/export share one geometry source; the dispatched transaction
   *  re-renders the canvas through the bridge. */
  #setPageSize(value?: string): void {
    const size = value ? PAPER_SIZES[value] : undefined;
    if (size) {
      this.#updateSectionGeometry({
        pageSize: {
          width: convertMillimetersToTwip(size[0]),
          height: convertMillimetersToTwip(size[1]),
        },
      });
    }
  }

  /** Apply orientation (portrait/landscape) — writes orientation onto
   *  page.size, deep-merged with the current (or engine-default) size so the
   *  projection can swap edges for landscape. */
  #setOrientation(value?: string): void {
    if (!value) return;
    const cur = (
      this.editor?.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions } | undefined
    )?.sectionProperties?.pageSize;
    const size =
      cur && typeof cur.width === "number" && typeof cur.height === "number"
        ? cur
        : { width: sectionPageSizeDefaults.WIDTH, height: sectionPageSizeDefaults.HEIGHT };
    this.#updateSectionGeometry({
      pageSize: { ...size, orientation: value as "portrait" | "landscape" },
    });
  }

  /** Apply a margin preset (normal/narrow/…) — writes the margins into the
   *  document-model sectionProperties so a page-setup change actually
   *  re-lays-out (the transaction re-renders the canvas). */
  #setMargins(value?: string): void {
    if (value && MARGINS[value]) {
      this.#updateSectionGeometry({ pageMargin: marginTwipsFromCss(MARGINS[value]) });
    }
  }

  /** The doc position carrying the current section's sectPr — the first
   *  section-carrying paragraph at/after the caret (OOXML: its sectPr ends
   *  that section), or null when the caret sits in the final section (the
   *  sectPr is body-level on doc.attrs). */
  #sectionSectPrPos(): number | null {
    const editor = this.editor;
    if (!editor) return null;
    const from = editor.state.selection.from;
    let targetPos: number | null = null;
    editor.state.doc.descendants((node, nodePos) => {
      if (targetPos != null) return true;
      // Paragraphs ending at/before the caret close earlier sections; a
      // paragraph CONTAINING the caret owns the current section (OOXML: its
      // sectPr ends that section, caret position included).
      if (nodePos + node.nodeSize <= from) return true;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        targetPos = nodePos;
        return false;
      }
      return true;
    });
    return targetPos;
  }

  /** Deep-merge a sectionProperties patch into the CURRENT section's sectPr and
   *  dispatch it — Word's "this section" semantics. The dispatched transaction
   *  re-renders every page of the canvas. */
  #updateSectionGeometry(patch: SectionPropertiesOptions): void {
    const editor = this.editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const targetPos = this.#sectionSectPrPos();
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

  /** Toggle a slot-visibility flag (titlePage / evenAndOddHeaders) on the
   *  current section's sectPr (Word's Different First Page / Odd & Even
   *  Pages). The transaction re-renders; the furniture projection picks the
   *  flag up and the page pattern (first/even slots) follows. */
  #toggleSectionFlag(flag: "titlePage" | "evenAndOddHeaders"): void {
    const editor = this.editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const flip = (cur: SectionPropertiesOptions | undefined): SectionPropertiesOptions => ({
      ...cur,
      [flag]: !(cur as unknown as Record<string, unknown> | undefined)?.[flag],
    });
    const targetPos = this.#sectionSectPrPos();
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        const cur = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        tr.setNodeMarkup(targetPos, undefined, { ...node.attrs, sectionProperties: flip(cur) });
      }
    } else {
      const cur = (doc.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
      tr.setDocAttribute("sectionProperties", flip(cur));
    }
    editor.view.dispatch(tr);
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

  /** Runtime `section-properties` change: deep-merge into the body section's
   *  sectPr (a default doc is single-section); the dispatched transaction
   *  re-renders the canvas with the new geometry. */
  #applySectionPropertiesAttr(): void {
    const editor = this.editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().sectionProperties;
    if (!parsed) return;
    const cur = (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("sectionProperties", mergeSectionProperties(cur, parsed)),
    );
  }

  /** Runtime `styles` change: replace doc.attrs.styles (the style library
   *  re-renders through the layout pipeline and the Styles gallery). */
  #applyStylesAttr(): void {
    const editor = this.editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().styles;
    if (parsed === undefined) return;
    editor.view.dispatch(editor.state.tr.setDocAttribute("styles", parsed));
    this.#renderChrome();
  }

  /** Apply a zoom level (percent, clamped 10–500) to the page stage and
   *  refresh the status bar. The stage sizes its slots to the scaled page
   *  directly (no CSS zoom — bitmaps stay 1:1 with screen pixels at every
   *  level). Idempotent (no-op on no change) and dispatches
   *  `docen:zoom-change` on a real flip — so the host, status-bar slider, and
   *  external listeners stay in sync through one funnel (Office
   *  `Office.Document.zoom.set` equivalent). */
  #setZoom(pct: number): void {
    const next = Math.max(10, Math.min(500, Math.round(pct)));
    if (next === this.#zoom) return;
    this.#zoom = next;
    this.#stage?.setZoom(next);
    // The frames resized under the overlays — re-place them at the new scale.
    this.#bridge?.replaceOverlays();
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
   *  to fill the document-area width (layout px at 100%). */
  #zoomPreset(preset: string): void {
    if (/^\d+$/.test(preset)) return this.#setZoom(Number(preset));
    if (preset !== "page-width") return;
    const area = this.shadowRoot?.querySelector("docen-document-area");
    if (!area || !this.#flow) return;
    this.#setZoom((area.clientWidth / this.#flow.pageWidthPx) * 100);
  }

  /** Refresh the status bar to mirror Word's bottom row: the left cluster is
   *  the caret's section, then "Page X of Y", then the word count; the right
   *  cluster is the zoom slider value + percent. Runs on every transaction
   *  (caret moves, a re-render changes the page count) and on zoom / locale
   *  change. The word count is cached by doc nodeSize so caret moves skip
   *  re-walking the full document. */
  #updateStatus(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const bar = root.querySelector<HTMLElement>("docen-status-bar");
    const editor = this.editor;
    const page = editor ? (this.#bridge?.pageOf(editor.state.selection.from) ?? -1) + 1 : 0;
    const total = this.#pages.length;
    // The caret's section: the section its page belongs to (1-based).
    const section = page > 0 ? (this.#sectionOfPage[page - 1] ?? 0) + 1 : 1;
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

  /** Word Count (Review tab) — compute the document statistics (the status
   *  bar's words source + grapheme character counts + layout line total) and
   *  hand them to the dialog as one JSON attribute. */
  #showWordCount(): void {
    const editor = this.editor;
    const dialog = this.shadowRoot?.querySelector("docen-word-count-dialog") as
      | (HTMLElement & { stats?: string; show(): void })
      | undefined;
    if (!editor || !dialog) return;
    const cc = editor.storage.characterCount as { words?: () => number } | undefined;
    const text = editor.state.doc.textContent;
    let paragraphs = 0;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "paragraph") paragraphs++;
      return true;
    });
    let lines = 0;
    for (const page of this.#pages) {
      for (const item of page.items) {
        if (item.block.kind === "paragraph") lines += item.block.lines.length;
      }
    }
    dialog.stats = JSON.stringify({
      pages: this.#pages.length,
      words: cc?.words?.() ?? 0,
      charsWithSpaces: textCounter(text),
      charsNoSpaces: textCounter(text.replace(/\s+/g, "")),
      paragraphs,
      lines,
    });
    dialog.show();
  }

  /** Symbol dialog Insert → drop the picked character at the caret (the
   *  dialog stays open, Word-style, so several symbols can go in a row). */
  readonly #onSymbolInsert = (event: CustomEvent<{ char?: string }>): void => {
    const char = event.detail?.char;
    if (!char) return;
    this.#bridge?.focus();
    this.editor?.commands.insertContent(char);
  };

  /** The next free bookmark id — one past the highest id already carried by
   *  any bookmarkStart/bookmarkEnd passthrough in the document (OOXML marks
   *  pairs by a document-unique integer). */
  #nextBookmarkId(): number {
    let max = -1;
    const scan = (node: JSONContent): void => {
      for (const child of node.content ?? []) {
        if (child.type === "inlinePassthrough" || child.type === "passthrough") {
          try {
            const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
              bookmarkStart?: { id?: number };
              bookmarkEnd?: { id?: number };
            };
            for (const id of [data.bookmarkStart?.id, data.bookmarkEnd?.id]) {
              if (typeof id === "number" && id > max) max = id;
            }
          } catch {
            // opaque verbatim blobs without bookmark data — skip
          }
        }
        scan(child);
      }
    };
    scan(this.editor!.getJSON());
    return max + 1;
  }

  /** Insert Bookmark — prompt for a name (Word's rules: starts with a letter
   *  or CJK char, no spaces), then wrap the selection with a
   *  bookmarkStart/bookmarkEnd passthrough pair in one transaction (the start
   *  goes before `from`, the end after `to` — +1 shifts past the start atom
   *  the first step added). Round-trips verbatim through DOCX. */
  #insertBookmark(): void {
    const editor = this.editor;
    if (!editor) return;
    const name = window.prompt(t("bookmark.prompt", this))?.trim();
    if (name == null) return;
    if (!/^[A-Za-z一-鿿぀-ヿ][^\s]*$/.test(name) || name.length > 40) {
      window.alert(t("bookmark.invalid", this));
      return;
    }
    const id = this.#nextBookmarkId();
    const seed = (data: object): JSONContent =>
      ({
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify(data) },
      }) as JSONContent;
    const { from, to } = editor.state.selection;
    const start = editor.schema.nodeFromJSON(seed({ bookmarkStart: { id, name } }));
    const end = editor.schema.nodeFromJSON(seed({ bookmarkEnd: { id } }));
    editor.view.dispatch(editor.state.tr.insert(from, start).insert(to + 1, end));
  }

  /** Insert → Link: prompt for the address (pre-filled with the selection's
   *  existing link, Word's edit-an-existing-hyperlink behavior), then either
   *  mark the selected text or insert fresh display text. An empty address on
   *  an existing link removes it (Word's "remove hyperlink"). `#name`
   *  addresses are bookmark anchors (in-page jumps); bare hosts gain the
   *  https scheme (Word's auto-complete). */
  #insertLink(): void {
    const editor = this.editor;
    if (!editor) return;
    // The link mark riding the selection, if any (extendMarkRange snaps the
    // range to the whole mark, so a caret inside a link edits the whole one).
    const existing = editor.getAttributes("link").href as string | undefined;
    const raw = window.prompt(t("link.prompt", this), existing ?? "https://")?.trim();
    if (raw == null) return;
    if (raw === "") {
      if (existing) editor.chain().focus().extendMarkRange("link").unsetLink().run();
      return;
    }
    const href = raw.startsWith("#") || /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
    const { empty } = editor.state.selection;
    // Word stamps inserted hyperlink runs with the "Hyperlink" character style
    // — that style (not the w:hyperlink element) paints links blue — so every
    // insert path here stamps it in the same transaction.
    if (!empty) {
      editor
        .chain()
        .focus()
        .extendMarkRange("link")
        .setLink({ href })
        .setMark("textStyle", { style: "Hyperlink" })
        .run();
      return;
    }
    // No selection: Word asks for display text and inserts it marked.
    const text = window
      .prompt(t("link.text.prompt", this), raw.replace(/^https?:\/\//, ""))
      ?.trim();
    if (!text) return;
    editor
      .chain()
      .focus()
      .insertContent({
        type: "text",
        text,
        marks: [
          { type: "link", attrs: { href, target: href.startsWith("#") ? null : "_blank" } },
          { type: "textStyle", attrs: { style: "Hyperlink" } },
        ],
      })
      .run();
  }

  /** References → Next Footnote: place the caret on the next
   *  footnote/endnote reference after the selection (document order); none is
   *  a no-op (Word steps through its notes without wrapping). */
  #jumpNextNote(): void {
    const editor = this.editor;
    if (!editor) return;
    const { from } = editor.state.selection;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (target != null) return false;
      if (pos <= from || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as Record<string, unknown>;
        if ("footnoteReference" in data || "endnoteReference" in data) target = pos;
      } catch {
        // opaque verbatim blob — not a note reference
      }
    });
    // After the leaf atom (pos is its left edge) so the caret sits past it.
    if (target != null) this.#setTextSelection(target + 1);
  }

  /** One inlinePassthrough comment atom (see #insertComment) → its marker
   *  kind and comment id; non-comment atoms yield null. Accepts both a JSON
   *  atom (type is the string name) and a PM node from doc.descendants
   *  (type is the NodeType — read its name). */
  static commentMarkerOf(
    child: JSONContent | { type?: { name?: string }; attrs?: { data?: string } },
  ): { id: number; kind: "start" | "end" | "reference" } | null {
    const type = child.type as string | { name?: string } | undefined;
    const typeName = typeof type === "string" ? type : type?.name;
    if (typeName !== "inlinePassthrough") return null;
    try {
      const data = JSON.parse(String((child.attrs as { data?: string } | undefined)?.data)) as {
        commentRangeStart?: { id?: number };
        commentRangeEnd?: { id?: number };
        commentReference?: number;
      };
      if (typeof data.commentRangeStart?.id === "number")
        return { id: data.commentRangeStart.id, kind: "start" };
      if (typeof data.commentRangeEnd?.id === "number")
        return { id: data.commentRangeEnd.id, kind: "end" };
      if (typeof data.commentReference === "number")
        return { id: data.commentReference, kind: "reference" };
    } catch {
      // Malformed passthrough data — not a comment marker.
    }
    return null;
  }

  /** The comment whose range covers the selection (a marker pair bracketing
   *  it in document order — markers may sit in earlier paragraphs), lowest id
   *  first; null when the selection touches no comment. */
  #activeCommentId(): number | null {
    const editor = this.editor;
    if (!editor) return null;
    const { from, to } = editor.state.selection;
    const opened = new Map<number, number>();
    const covering = new Set<number>();
    editor.state.doc.descendants((child, pos) => {
      const marker = DocenDocument.commentMarkerOf(child);
      if (!marker) return;
      if (marker.kind === "start") opened.set(marker.id, pos);
      else if (marker.kind === "end") {
        const start = opened.get(marker.id);
        if (start != null && start < to && pos > from) covering.add(marker.id);
      }
    });
    return covering.size > 0 ? Math.min(...covering) : null;
  }

  /** Review → Edit Comment: open the comments pane — editing happens inline
   *  on the card (Word edits comments in the sidebar, not a dialog). */
  #editComment(): void {
    this.showTaskpane("comments");
  }

  /** Review → Delete: remove the covering comment's marker/reference atoms
   *  (descending positions keep the earlier offsets valid) and its
   *  documentExtras entry. */
  #deleteComment(): void {
    const id = this.#activeCommentId();
    if (id == null) return;
    this.#deleteCommentById(id);
  }

  /** Remove a comment's marker/reference atoms (descending positions keep the
   *  earlier offsets valid) and its documentExtras entry — shared by the
   *  ribbon's Delete Comment and the pane's per-card delete. */
  #deleteCommentById(id: number): void {
    const editor = this.editor;
    if (!editor) return;
    const atoms: { pos: number; size: number }[] = [];
    editor.state.doc.descendants((child, pos) => {
      const marker = DocenDocument.commentMarkerOf(child);
      if (marker && marker.id === id) atoms.push({ pos, size: child.nodeSize });
    });
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Record<string, unknown>[] };
    };
    const tr = editor.state.tr;
    for (const { pos, size } of atoms.sort((a, b) => b.pos - a.pos)) tr.delete(pos, pos + size);
    tr.setDocAttribute("documentExtras", {
      ...docAttrs.documentExtras,
      comments: (docAttrs.documentExtras?.comments ?? []).filter((c) => Number(c.id) !== id),
    });
    editor.view.dispatch(tr);
  }

  /** Review → Previous/Next Comment: select the range of the comment before
   *  or after the selection (document order); no further comment in that
   *  direction is a no-op. */
  #jumpComment(direction: "previous" | "next"): void {
    const editor = this.editor;
    if (!editor) return;
    const ranges: { from: number; to: number }[] = [];
    const openStarts = new Map<number, number>();
    editor.state.doc.descendants((child, pos) => {
      const marker = DocenDocument.commentMarkerOf(child);
      if (!marker) return;
      if (marker.kind === "start") openStarts.set(marker.id, pos + child.nodeSize);
      else if (marker.kind === "end") {
        const start = openStarts.get(marker.id);
        if (start != null) ranges.push({ from: start, to: pos });
      }
    });
    ranges.sort((a, b) => a.from - b.from);
    const { from } = editor.state.selection;
    const target =
      direction === "next"
        ? ranges.find((r) => r.from > from)
        : ranges.findLast((r) => r.from < from);
    if (!target) return;
    editor.view.dispatch(
      editor.state.tr.setSelection(
        new TextSelection(
          editor.state.doc.resolve(target.from),
          editor.state.doc.resolve(target.to),
        ),
      ),
    );
  }

  /** Review → New Comment: open the comments pane's compose box on the
   *  current selection (Word's sidebar compose card); the text arrives via
   *  the `comment:create` event (#onCommentCreate commits it). */
  #insertComment(): void {
    const editor = this.editor;
    if (!editor) return;
    this.showTaskpane("comments");
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    // Spread would miss from/to — they're prototype getters on Selection.
    const { from, to } = editor.state.selection;
    this.#pendingCommentRange = { from, to };
    pane.setAttribute("draft", "");
  }

  /** comment:create → commit the pending comment: anchor the stored range the
   *  way Word does — a commentRangeStart/commentRangeEnd passthrough pair
   *  around it with a commentReference after — and append the structured
   *  content to doc.attrs.documentExtras.comments (word/comments.xml on
   *  export, the round-trip channel the parse side already fills). */
  readonly #onCommentCreate = (event: CustomEvent<{ text?: string }>): void => {
    const editor = this.editor;
    const text = event.detail?.text?.trim();
    const range = this.#pendingCommentRange;
    this.#pendingCommentRange = undefined;
    this.#commentsPaneEl()?.removeAttribute("draft");
    if (!editor || !text || !range) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Record<string, unknown>[] };
    };
    const comments = docAttrs.documentExtras?.comments ?? [];
    const id = comments.reduce((max, c) => Math.max(max, Number(c.id ?? 0)), -1) + 1;
    const seed = (data: object): JSONContent =>
      ({
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify(data) },
      }) as JSONContent;
    const start = editor.schema.nodeFromJSON(seed({ commentRangeStart: { id } }));
    const end = editor.schema.nodeFromJSON(seed({ commentRangeEnd: { id } }));
    const ref = editor.schema.nodeFromJSON(seed({ commentReference: id }));
    editor.view.dispatch(
      editor.state.tr
        .insert(range.from, start)
        .insert(range.to + 1, end)
        .insert(range.to + 2, ref)
        .setDocAttribute("documentExtras", {
          ...docAttrs.documentExtras,
          comments: [
            ...comments,
            {
              id,
              author: "Docen User",
              initials: "DU",
              date: new Date().toISOString(),
              children: [{ text }],
            },
          ],
        }),
    );
  };

  /** comment:cancel → drop the pending compose (Word's Cancel). */
  readonly #onCommentCancel = (): void => {
    this.#pendingCommentRange = undefined;
    this.#commentsPaneEl()?.removeAttribute("draft");
  };

  /** comment:select → select and scroll to the comment's range (Word scrolls
   *  the anchored text into view when its card is clicked). */
  readonly #onCommentSelect = (event: CustomEvent<{ id?: number }>): void => {
    const editor = this.editor;
    const id = event.detail?.id;
    if (!editor || id == null) return;
    const range = this.#commentRangeOf(id);
    if (!range) return;
    editor.view.dispatch(
      editor.state.tr.setSelection(
        new TextSelection(editor.state.doc.resolve(range.from), editor.state.doc.resolve(range.to)),
      ),
    );
    this.#bridge?.scrollIntoView(range.from);
  };

  /** comment:update → rewrite the comment's text (the sidebar's inline edit,
   *  replacing the old prompt-based Edit Comment for pane interactions). */
  readonly #onCommentUpdate = (event: CustomEvent<{ id?: number; text?: string }>): void => {
    const editor = this.editor;
    const id = event.detail?.id;
    const text = event.detail?.text?.trim();
    if (!editor || id == null || !text) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Array<{ id?: number; children?: Array<{ text?: string }> }> };
    };
    const comments = docAttrs.documentExtras?.comments ?? [];
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", {
        ...docAttrs.documentExtras,
        comments: comments.map((c) => (Number(c.id) === id ? { ...c, children: [{ text }] } : c)),
      }),
    );
  };

  /** comment:delete → remove the comment's marker/reference atoms and its
   *  documentExtras entry (same teardown as the ribbon's Delete Comment). */
  readonly #onCommentDelete = (event: CustomEvent<{ id?: number }>): void => {
    const id = event.detail?.id;
    if (id != null) this.#deleteCommentById(id);
  };

  /** The comments pane element (inside its task-pane), when mounted. */
  #commentsPaneEl(): (HTMLElement & { comments?: string; draft?: boolean }) | null {
    return this.shadowRoot?.querySelector("docen-comments-pane") ?? null;
  }

  /** The document order range a comment id covers (start marker through end
   *  marker), or null when its markers are gone. */
  #commentRangeOf(id: number): { from: number; to: number } | null {
    const editor = this.editor;
    if (!editor) return null;
    let from: number | null = null;
    let to: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      const marker = DocenDocument.commentMarkerOf(child);
      if (!marker || marker.id !== id) return;
      if (marker.kind === "start") from = pos + child.nodeSize;
      else if (marker.kind === "end") to = pos;
    });
    return from != null && to != null ? { from, to } : null;
  }

  /** Sync the comments pane's card list from the doc's documentExtras after
   *  every transaction (the pane is a pure view of the model). */
  #syncCommentsPane(): void {
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    const docAttrs = (this.editor?.state.doc.attrs ?? {}) as {
      documentExtras?: {
        comments?: Array<{
          id?: number;
          author?: string;
          initials?: string;
          date?: string;
          children?: Array<{ text?: string }>;
        }>;
      };
    };
    const cards = (docAttrs.documentExtras?.comments ?? []).map((c) => ({
      id: Number(c.id ?? 0),
      author: c.author ?? "",
      initials: c.initials ?? "",
      date: c.date ?? "",
      text: (c.children ?? []).map((r) => r.text ?? "").join(""),
    }));
    pane.comments = JSON.stringify(cards);
  }

  /** Insert → Footnote / Endnote (Word's References → Insert Footnote, the
   *  split's two items): prompt for the note text, drop a footnoteReference /
   *  endnoteReference atom at the caret and append the note's content to
   *  doc.attrs.documentExtras.footnotes/endnotes (word/footnotes.xml /
   *  word/endnotes.xml on export — the round-trip channel the parse side
   *  already fills). The note body opens with the noteRef mark run (Word's
   *  note-number placeholder); the bare reference atom picks up the built-in
   *  FootnoteReference/EndnoteReference style on export (the canvas paints
   *  endnote references lowercase Roman, Word's endnote default numFmt).
   *  A document carrying explicit content types needs the matching override
   *  added too — the packer gates the part on it whenever contentTypes exists. */
  #insertNote(kind: "footnote" | "endnote"): void {
    const editor = this.editor;
    if (!editor) return;
    const text = window.prompt(t("footnote.prompt", this))?.trim();
    if (!text) return;
    const channel = `${kind}s` as "footnotes" | "endnotes";
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: {
        footnotes?: Array<{ id?: number }>;
        endnotes?: Array<{ id?: number }>;
        contentTypes?: {
          overrides?: Array<{ partName?: string; contentType?: string }>;
        };
      };
    };
    const extras = docAttrs.documentExtras ?? {};
    const notes = extras[channel] ?? [];
    const id = notes.reduce((max, n) => Math.max(max, Number(n.id ?? 0)), 0) + 1;
    const Note = kind === "footnote" ? "FootnoteText" : "EndnoteText";
    const ref = editor.schema.nodeFromJSON({
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify({ [`${kind}Reference`]: id }) },
    } as JSONContent);
    const documentExtras = {
      ...extras,
      [channel]: [
        ...notes,
        {
          id,
          // Word's canonical note body: one *Text paragraph holding the
          // note-number mark run (the bare *Ref atom — export wraps it in the
          // *Reference rStyle) followed by the note text inline.
          children: [
            {
              style: Note,
              children: [{ [`${kind}Ref`]: true }, { text }],
            },
          ],
        },
      ],
    };
    const partName = `/word/${channel}.xml`;
    if (
      extras.contentTypes &&
      !extras.contentTypes.overrides?.some((o) => o.partName === partName)
    ) {
      documentExtras.contentTypes = {
        ...extras.contentTypes,
        overrides: [
          ...(extras.contentTypes.overrides ?? []),
          {
            partName,
            contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${channel}+xml`,
          },
        ],
      };
    }
    editor.view.dispatch(
      editor.state.tr
        .insert(editor.state.selection.from, ref)
        .setDocAttribute("documentExtras", documentExtras as typeof extras),
    );
  }

  /** Design → Page Color: write the doc-level page background
   *  (doc.attrs.background → w:background on export; the stage paints it as
   *  the page frame color). "none" clears it (Word's No Color); a bare hex is
   *  the standard/custom swatch path; a theme-semantic pick carries its
   *  themeColor/tint/shade through so Word re-resolves on theme change. */
  #setPageColor(
    value?: string | { themeColor: string; val: string; themeTint?: string; themeShade?: string },
  ): void {
    const editor = this.editor;
    if (!editor) return;
    const background =
      value == null || value === "none"
        ? null
        : typeof value === "string"
          ? { color: value }
          : {
              color: value.val,
              themeColor: value.themeColor,
              ...(value.themeTint ? { themeTint: value.themeTint } : {}),
              ...(value.themeShade ? { themeShade: value.themeShade } : {}),
            };
    editor.view.dispatch(editor.state.tr.setDocAttribute("background", background));
  }

  /** Insert → Text Box / Shapes: a standalone wps shape run, floating
   *  wrap-none and centered on the page (Word's insertion behavior). The
   *  text box carries Word's plain look — white fill, accent-1 hairline —
   *  and an editable empty body (the PM `content`); a gallery shape carries
   *  its preset geometry with the accent fill instead. */
  #insertShape(preset: string | undefined): void {
    const editor = this.editor;
    if (!editor) return;
    const geometry: Record<string, unknown> = {
      // Word's plain text box default: 2" × 1.2".
      transformation: { width: 1828800, height: 1097280 },
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", align: "center" },
        wrap: { type: "none" },
      },
    };
    if (preset) {
      geometry.presetGeometry = { preset };
      // The theme's accent-1 pair (fill + its darkened outline) — the same
      // look Word gives a fresh shape; the projection paints flat hex.
      geometry.fill = { type: "solid", color: "4472C4" };
      geometry.outline = { color: "2F528F", width: 12700 };
    } else {
      geometry.fill = { type: "solid", color: "FFFFFF" };
      geometry.outline = { color: "4472C4", width: 12700 };
    }
    editor.commands.insertContentAt(editor.state.selection.from, {
      type: "wpsShape",
      attrs: { wpsShape: geometry },
      content: [{ type: "paragraph" }],
    } as JSONContent);
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
    // Word Count (ribbon Review → Proofing) → the statistics dialog.
    if (name === "word-count") {
      this.#showWordCount();
      return;
    }
    // Page setup actions write sectionProperties; the transaction re-renders.
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
    const editor = this.editor;
    if (!editor) return;
    // Edit / View mode — toggle the editor's editable state (tab-row "Editing"
    // menu); then re-stamp the menu so its label + checked item follow.
    if (name === "edit-mode") {
      editor.setEditable(value !== "view");
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
    // Formatting marks toggle — canvas-side marks are a later milestone; the
    // host [show-marks] attribute stays the source of truth.
    if (name === "show-marks") {
      this.setShowMarks(!this.getShowMarks());
      return;
    }
    // TOC insert/update — commands take the bridge's pageOf (entry page
    // numbers come from the canvas caret map; 0-based → Word's 1-based) and
    // the content-width tab stop. Inserting repaginates, so insert re-runs
    // the update once the fresh layout lands (Word's insert-then-update-
    // fields behavior).
    if (name === "toc" || name === "update-toc") {
      const pageOf = (pos: number): number | null => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const tabPositionTw = this.#flow
        ? Math.round(this.#flow.contentWidthPx / twipToPx(1))
        : undefined;
      const ran = editor.commands[name](pageOf, tabPositionTw);
      if (name === "toc" && ran) {
        // Frame N re-flows (the bridge's raf-merged onDoc), frame N+1 the
        // caret map carries the post-insert pagination.
        requestAnimationFrame(() =>
          requestAnimationFrame(() => editor.commands["update-toc"](pageOf, tabPositionTw)),
        );
      }
      return;
    }
    // Header/Footer — the split's main action opens the story on the caret's
    // page; the drop-down carries remove + the slot-visibility flags.
    if (name === "header" || name === "footer") {
      if (value === "title-page" || value === "odd-even") {
        this.#toggleSectionFlag(value === "title-page" ? "titlePage" : "evenAndOddHeaders");
        return;
      }
      if (value === "remove-header" || value === "remove-footer") {
        this.#removeStory(name);
        return;
      }
      const page = this.#bridge?.pageOf(editor.state.selection.from);
      if (page != null) this.#bridge?.enterStory(name, page);
      return;
    }
    // Page Number — seed a PAGE field at the chosen story's end (Word's
    // default is bottom of page); the story stays open so the user can
    // adjust, and the normal exit persists the slots.
    if (name === "page-number") {
      if (value === "remove-numbers") {
        this.#removePageNumbers();
        return;
      }
      const page = this.#bridge?.pageOf(editor.state.selection.from);
      const seed: JSONContent = {
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify({ simpleField: { instruction: "PAGE" } }) },
      };
      if (page != null) {
        this.#bridge?.enterStory(value === "top" ? "header" : "footer", page, seed);
      }
      return;
    }
    // Symbol — open the character grid dialog; the insertion arrives via the
    // dialog's symbol:insert event (it stays open for several inserts).
    if (name === "symbol") {
      (this.shadowRoot?.querySelector("docen-symbol-dialog") as { show(): void } | null)?.show();
      return;
    }
    // Bookmark — prompt for a name and wrap the selection with a
    // bookmarkStart/bookmarkEnd pair (Word's Insert → Bookmark).
    if (name === "bookmark") {
      this.#insertBookmark();
      return;
    }
    // Footnote / Endnote — prompt for the note text, reference the caret and
    // append the note body (Word's References → Insert Footnote; the split's
    // endnote item shares the event, and Next Footnote steps references).
    if (name === "insert-footnote") {
      if (value === "endnote") this.#insertNote("endnote");
      else if (value === "next") this.#jumpNextNote();
      else this.#insertNote("footnote");
      return;
    }
    // Page Color — write/clear the doc-level w:background from the palette
    // (Word's Design → Page Color).
    if (name === "page-color") {
      this.#setPageColor(
        value as
          | string
          | { themeColor: string; val: string; themeTint?: string; themeShade?: string },
      );
      return;
    }
    // Link — prompt for an address and mark the selection (or insert fresh
    // display text when the selection is empty).
    if (name === "link") {
      this.#insertLink();
      return;
    }
    // New Comment — anchor the selection with a Word comment; Edit/Delete
    // operate on the comment covering the selection. The Review tab's large
    // button and the trailing toolbar comment share the "comment" event.
    if (name === "new-comment" || name === "comment") {
      this.#insertComment();
      return;
    }
    if (name === "edit-comment") {
      this.#editComment();
      return;
    }
    if (name === "delete-comment") {
      this.#deleteComment();
      return;
    }
    if (name === "previous-comment") {
      this.#jumpComment("previous");
      return;
    }
    if (name === "next-comment") {
      this.#jumpComment("next");
      return;
    }
    // Review → Show Comments: toggle the comments pane (Word's sidebar).
    if (name === "show-comments") {
      this.#setTaskpane("comments", !this.getTaskpaneState("comments"));
      return;
    }
    // Text Box / Shapes — insert a floating wps shape run (Shapes reads its
    // preset from the gallery item's value; the text box has no preset).
    if (name === "text-box" || name === "shapes") {
      this.#insertShape(name === "shapes" ? (value ?? "rect") : undefined);
      return;
    }
    // Clipboard — the selection is canvas-rendered (no DOM editor selection),
    // so copy/cut read the doc range and write the clipboard directly.
    if (name === "copy" || name === "cut") {
      void this.#copySelection(name === "cut");
      return;
    }
    if (name === "paste") {
      void this.#paste(value === "keep-text-only");
      return;
    }
    // Editing → Select: selectAll() spans the whole document.
    if (name === "select") {
      this.#select(value);
      return;
    }
    // Format Painter — toggle capture/apply of the current run's marks.
    if (name === "format-painter") {
      this.#toggleFormatPainter();
      return;
    }
    // Built-in commands route to editor.commands.<event>(value) —
    // DocumentCommands registers every ribbon event as a native Tiptap command.
    // A user add-in overrides one by contributing a Tiptap extension whose
    // addCommands redefines the same name (Tiptap's native override mechanism).
    const commands = editor.commands as unknown as Record<string, (value?: string) => unknown>;
    const cmd = commands[name];
    if (typeof cmd === "function") {
      cmd(value);
      return;
    }
    // Not a Tiptap command — route to the first add-in that declares it. This
    // covers non-Tiptap actions contributed by external add-ins (e.g. a Help
    // button that opens a URL) that Tiptap can't express.
    this.dispatchCommand(name, value);
  };

  /** Copy/cut the current selection to the system clipboard as plain text;
   *  cut also deletes the range. */
  async #copySelection(cut: boolean): Promise<void> {
    const editor = this.editor;
    if (!editor) return;
    const { from, to } = editor.state.selection;
    if (from === to) return;
    const text = editor.state.doc.textBetween(from, to, "\n");
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // Clipboard write may be denied (permissions/policy) — still cut.
    }
    if (cut) {
      editor.view.dispatch(editor.state.tr.deleteSelection());
    }
  }

  /** Menu items and the auto-save switch carry their action in `data-event`. */
  readonly #onChange = (event: Event): void => {
    const name = (event.target as HTMLElement)?.dataset?.event;
    if (!name) return;
    switch (name) {
      case "open":
        // Host can take over via docen:open (preventDefault); else open the
        // picker — #onFileChange auto-detects docx/md from the extension.
        if (!this.#emitCancelable("docen:open")) this.#pickFile();
        break;
      case "save-as":
        if (!this.#emitCancelable("docen:save-as", { format: "docx" })) void this.#saveAs("docx");
        break;
      case "save-as-markdown":
        if (!this.#emitCancelable("docen:save-as", { format: "markdown" }))
          void this.#saveAs("markdown");
        break;
      case "print":
        if (!this.#emitCancelable("docen:print")) this.#print();
        break;
      case "new":
        // No built-in "new" — always hand to the host (docen:new).
        this.#emitCancelable("docen:new");
        break;
      case "options": {
        // Filename menu → open the Options dialog (UI language + theme).
        const optionsEl = this.shadowRoot?.querySelector("docen-options-dialog");
        if (optionsEl) {
          optionsEl.setAttribute("locale", this.lang || document.documentElement.lang || "zh-CN");
          optionsEl.setAttribute("theme", this.theme ?? "light");
          (optionsEl as unknown as { show?: () => void }).show?.();
        }
        break;
      }
      // autosave: skeleton — wired when that feature lands.
    }
  };

  /** Forward this host's `lang` attribute to the internal <docen-workspace>
   *  and notify locale observers. Called on connect and whenever `lang`
   *  mutates (via #langObserver). The workspace is the resolveLang scope, so
   *  forwarding is what makes <docen-document lang> reach child components
   *  across the shadow boundary. */
  #syncLang(): void {
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    const lang = this.lang;
    if (lang) workspace?.setAttribute("lang", lang);
    else workspace?.removeAttribute("lang");
    notifyLocaleChange();
  }

  readonly #onLangChange = (event: Event): void => {
    const lang = (event as CustomEvent<{ lang: string }>).detail?.lang;
    // Set on the host (not <html lang>) so locale is per-instance; the
    // #langObserver forwards to the workspace and notifies observers.
    if (lang) {
      this.setAttribute("lang", lang);
      this.#emitLangChange(lang);
    }
  };

  /** Options dialog 确定 — commit the UI language + theme. */
  readonly #onOptionsOk = (event: Event): void => {
    const { lang, theme } = (event as CustomEvent<{ lang?: string; theme?: string }>).detail ?? {};
    if (lang && this.getAttribute("lang") !== lang) {
      this.setAttribute("lang", lang);
      this.#emitLangChange(lang);
    }
    if (theme && this.getAttribute("theme") !== theme) {
      this.setAttribute("theme", theme);
      this.#emitThemeChange(theme);
    }
  };

  /** Notify external listeners (framework wrappers like @docen/vue) when the
   *  locale changes from inside the host — status-bar toggle or Options OK.
   *  External `lang` writes (e.g. a Vue prop) set the attribute directly and
   *  don't route through here, so there's no echo cycle. */
  #emitLangChange(lang: string): void {
    this.dispatchEvent(
      new CustomEvent("docen:lang-change", { bubbles: true, composed: true, detail: { lang } }),
    );
  }

  /** Notify external listeners (framework wrappers like @docen/vue) when the
   *  theme changes from inside the host — Options OK. External `theme` writes
   *  (e.g. a Vue prop) set the attribute directly and don't route through
   *  here, so there's no echo cycle. */
  #emitThemeChange(theme: string): void {
    this.dispatchEvent(
      new CustomEvent("docen:theme-change", { bubbles: true, composed: true, detail: { theme } }),
    );
  }

  /** Open the OS file picker. The accept filter on the input element covers
   *  .docx/.md/.markdown; #onFileChange routes the chosen file by extension
   *  via open(). */
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
   *  canvas renders the natural size, and prepareImages fills them on DOCX
   *  export. */
  readonly #onImageChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      this.#bridge?.focus();
      this.editor?.commands.insertContent({ type: "image", attrs: { src: reader.result } });
    };
    reader.readAsDataURL(file);
  };

  /** Save the document in the given format via the native Save As dialog
   *  (showSaveFilePicker) when available so the user picks the location and name;
   *  falls back to a plain download otherwise. The header filename is updated to
   *  match the saved name. */
  async #saveAs(format: "docx" | "markdown" = "docx"): Promise<void> {
    const cfg = SAVE_FORMATS[format];
    // saveDOCX returns a buffer; Markdown returns a string.
    const data = format === "docx" ? await this.saveDOCX() : this.saveMarkdown();
    const blob = new Blob([data as BlobPart], { type: cfg.mime });
    // Re-stamp the extension so a .docx opened then saved as Markdown does not
    // keep its .docx name.
    const baseName = (this.getAttribute("filename")?.trim() || t("header.doc-name", this)).replace(
      /\.(docx|md|markdown|txt)$/i,
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

  /** Print only the document pages — never the ribbon/chrome. Each page
   *  canvas rasterizes into a hidden print-only iframe (one image per page at
   *  the page's true paper size, @page margin 0), so the browser's print
   *  dialog receives exactly the paginated document, like Word's print
   *  output. */
  #print(): void {
    const shots = this.#stage?.printSnapshots() ?? [];
    if (shots.length === 0) return;
    const first = shots[0]!;
    const frame = document.createElement("iframe");
    Object.assign(frame.style, {
      position: "fixed",
      right: "0",
      bottom: "0",
      width: "0",
      height: "0",
      border: "0",
    });
    document.body.append(frame);
    const doc = frame.contentDocument!;
    doc.open();
    doc.write(`<!doctype html><html><head><title>${this.getAttribute("filename") ?? "Document"}</title><style>
      @page { size: ${first.width / 96}in ${first.height / 96}in; margin: 0; }
      html, body { margin: 0; }
      img { display: block; width: 100%; }
      .pg { page-break-after: always; break-after: page; }
      .pg:last-child { page-break-after: auto; break-after: auto; }
    </style></head><body>`);
    for (const s of shots) doc.write(`<div class="pg"><img src="${s.url}"></div>`);
    doc.write("</body></html>");
    doc.close();
    frame.onload = () => {
      const win = frame.contentWindow;
      if (!win) return;
      const cleanup = (): void => frame.remove();
      win.addEventListener("afterprint", cleanup, { once: true });
      win.focus();
      win.print();
      // afterprint can lag behind the dialog closing — sweep after a grace.
      setTimeout(cleanup, 30_000);
    };
  }

  /** Common load path for openDOCX/openMarkdown: adopt a filename, replace the
   *  whole doc node. The #loadDoc wake-up transaction re-renders the canvas
   *  through the bridge. */
  #applyOpenedJSON(json: JSONContent, filename?: string): void {
    if (filename) this.setAttribute("filename", filename);
    this.#loadDoc(json);
  }

  /** Load a file into the editor, auto-detecting its format from the extension
   *  (.docx → DOCX, .md/.markdown → Markdown). This is the single entry point
   *  the filename-menu "Open…" uses; openDOCX/openMarkdown remain for when the
   *  caller already knows the format (e.g. loading a server-fetched docx buffer
   *  that has no filename). Throws on an unrecognized extension. */
  async open(file: File): Promise<void> {
    const format = detectOpenFormat(file);
    if (format === "docx") return this.openDOCX(file);
    return this.openMarkdown(file);
  }

  /** Load a .docx into the editor from a File or a buffer (ArrayBuffer /
   *  Uint8Array). A File also adopts its name as the filename; a bare buffer
   *  carries no name. parseDOCX is synchronous, but this is async so a File's
   *  bytes can be awaited. Large files report progress in the status bar
   *  (Word's bottom-row "Opening…"): streaming the bytes gives a real
   *  percentage; parsing and first layout show an indeterminate bar. */
  async openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void> {
    const name = input instanceof File ? input.name : undefined;
    this.#setProgress(t("status.opening", this).replace("{name}", name ?? "DOCX"));
    try {
      const buffer = input instanceof File ? await this.#readBytesProgressive(input) : input;
      // parseDOCX blocks the main thread — yield two frames so the read-stage
      // bar paints before it freezes, then swap to the indeterminate parse bar.
      await this.#nextFrame();
      this.#setProgress(t("status.parsing", this));
      const json = parseDOCX(buffer);
      this.#applyOpenedJSON(json, name);
      await this.#nextFrame();
      this.#setProgress();
    } catch (err) {
      this.#setProgress();
      throw err;
    }
  }

  /** Load a Markdown file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. */
  async openMarkdown(input: File | string): Promise<void> {
    const name = typeof input === "string" ? undefined : input.name;
    this.#setProgress(t("status.opening", this).replace("{name}", name ?? "Markdown"));
    try {
      const text = typeof input === "string" ? input : await input.text();
      await this.#nextFrame();
      this.#setProgress(t("status.parsing", this));
      this.#applyOpenedJSON(parseMarkdown(text), name);
      await this.#nextFrame();
      this.#setProgress();
    } catch (err) {
      this.#setProgress();
      throw err;
    }
  }

  /** Show (label + optional 0-100 value; absent = indeterminate) or clear the
   *  status bar's open-progress cluster. */
  #setProgress(label?: string, value?: number): void {
    const bar = this.shadowRoot?.querySelector("docen-status-bar");
    if (!bar) return;
    if (label == null) bar.removeAttribute("progress");
    else bar.setAttribute("progress", JSON.stringify({ label, value }));
  }

  /** Read a File's bytes through its stream so the progress bar tracks real
   *  bytes (File.arrayBuffer() is opaque). */
  async #readBytesProgressive(file: File): Promise<ArrayBuffer> {
    const total = file.size || 1;
    const reader = file.stream().getReader();
    const chunks: Uint8Array[] = [];
    let loaded = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done || !value) break;
      chunks.push(value);
      loaded += value.byteLength;
      this.#setProgress(
        t("status.opening", this).replace("{name}", file.name),
        5 + Math.round((loaded / total) * 35),
      );
    }
    const out = new Uint8Array(loaded);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return out.buffer;
  }

  /** Two rAFs — enough for the current progress state to paint before a
   *  synchronous block (parseDOCX) freezes the frame. */
  #nextFrame(): Promise<void> {
    return new Promise((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
  }

  /** Serialize the current document to a DOCX buffer. */
  async saveDOCX(): Promise<Uint8Array> {
    const buffer = await generateDOCX(this.getJSON());
    return buffer as unknown as Uint8Array;
  }

  /** Serialize the current document to a Markdown string. */
  saveMarkdown(): string {
    return generateMarkdown(this.getJSON());
  }

  /** Current document as Tiptap JSON. Cached — recomputed only after a doc
   *  change (see #onTransaction). */
  getJSON(): JSONContent {
    const editor = this.editor;
    if (!editor) return {} as JSONContent;
    if (this.#jsonDirty || this.#cachedJSON === undefined) {
      this.#cachedJSON = editor.getJSON();
      this.#jsonDirty = false;
    }
    return this.#cachedJSON;
  }

  /** Replace the document with Tiptap JSON. */
  setJSON(json: JSONContent): void {
    // A hand-built JSON (not from parseDOCX) lacks office-open's document-level
    // schema defaults — doc.attrs.styles (docDefaults body font/size/spacing)
    // and doc.attrs.sectionProperties (page size/margins/docGrid linePitch).
    // Without them the document has no body font, no page geometry, and no grid
    // for snapToGrid to pitch against. Normalize once on the way in; a doc that
    // already carries sectionProperties (a parseDOCX/getJSON round-trip) is a
    // no-op (normalizeDocument shallow-merges user attrs over defaults).
    if (!(json.attrs as { sectionProperties?: unknown } | undefined)?.sectionProperties) {
      json = normalizeDocument(json);
    }
    this.#loadDoc(json);
    this.#renderChrome();
  }

  /** Replace the whole doc node (content + doc-level attrs) via a fresh
   *  EditorState. Tiptap's setContent only swaps content and drops doc-level
   *  attrs; this carries them (styles/core/sectionProperties). updateState
   *  bypasses appendTransaction/onTransaction, so extensions that react to doc
   *  changes wouldn't wake — dispatch a docChanged tr (re-stamp the first
   *  block's attrs, a no-op visually) to trigger them: Outline re-reports the
   *  anchor list, and the bridge's raf-merged onDoc re-renders the canvas. */
  #loadDoc(doc: JSONContent): void {
    const editor = this.editor;
    if (!editor) return;
    // New document — invalidate the JSON cache.
    this.#jsonDirty = true;
    editor.view.updateState(
      EditorState.create({ doc: editor.schema.nodeFromJSON(doc), plugins: editor.state.plugins }),
    );
    // NOTE: no isDestroyed guard — the viewless editor's `isDestroyed` getter
    // defaults to true (it reads editorView, which element:null never sets).
    // updateState bypasses appendTransaction, so extensions that react to doc
    // changes wouldn't wake. Dispatch a docChanged tr to fire them. The tr
    // re-stamps the LAST leaf block's OWN attrs — a true no-op (same node,
    // same attrs) — so nothing is clobbered.
    const state = editor.state;
    // Last textblock/leaf block (deepest, rightmost) for the re-stamp — found
    // by descending the rightmost-child chain (O(depth)) instead of a full
    // nodesBetween scan (O(n)).
    const last = this.#lastMarkupTarget(state.doc);
    if (last) {
      // addToHistory:false — this re-stamp is an intentional no-op (same node,
      // same attrs) whose sole purpose is to fire appendTransaction (updateState
      // bypasses it). Left in history, it plants a no-op undo entry at the stack
      // bottom (undo returns true but changes nothing); excluding it keeps the
      // undo stack clean after load.
      editor.view.dispatch(
        state.tr.setNodeMarkup(last.pos, undefined, last.attrs).setMeta("addToHistory", false),
      );
    } else {
      // An empty document has no markup target — render directly.
      this.#renderDoc(editor.getJSON());
    }
  }

  /** Last textblock/leaf block (deepest, rightmost) for the #loadDoc re-stamp
   *  hack — the re-stamp target that fires the extension wake-up. Runs only on
   *  load (setJSON/openDOCX), not per edit, so the walk cost is amortized over
   *  the load itself. */
  #lastMarkupTarget(doc: import("@tiptap/pm/model").Node): {
    pos: number;
    attrs: Record<string, unknown>;
  } | null {
    let last: { pos: number; attrs: Record<string, unknown> } | null = null;
    doc.nodesBetween(0, doc.content.size, (node, pos) => {
      if (node.isText) return;
      if (node.isTextblock || node.isLeaf) {
        last = { pos, attrs: node.attrs as Record<string, unknown> };
      }
      // Don't descend into textblocks (their text isn't a markup target).
      return node.isTextblock ? false : undefined;
    });
    return last;
  }

  /** The underlying Tiptap editor (for advanced, direct control). */
  getEditor(): Editor | undefined {
    return this.editor;
  }

  /** Force a full canvas re-render now. */
  repaginate(): void {
    const editor = this.editor;
    if (editor) this.#renderDoc(editor.getJSON());
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

  /** Whether a task pane is currently open. Returns a boolean for convenience
   *  (callers want open/closed); the `docen:taskpane-visibility-change` event
   *  detail carries the string `VisibilityMode` to mirror `Office.VisibilityMode`. */
  getTaskpaneState(id: TaskPaneId): boolean {
    return !!this.#paneEl(id)?.open;
  }

  #paneEl(id: TaskPaneId): (HTMLElement & { open: boolean }) | null {
    // Panes are looked up by part, not position — two panes can share a side
    // (properties + comments both live on the end rail).
    const part =
      id === "navigation" ? "nav-pane" : id === "comments" ? "comments-pane" : "props-pane";
    return this.shadowRoot?.querySelector(`docen-task-pane[part="${part}"]`) as
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
   *  truth. Canvas-side mark rendering (¶ pilcrows, break dividers) is a later
   *  milestone — the attribute + event contract holds either way. */
  setShowMarks(on: boolean): void {
    if (this.hasAttribute("show-marks") === on) return;
    this.toggleAttribute("show-marks", on);
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
