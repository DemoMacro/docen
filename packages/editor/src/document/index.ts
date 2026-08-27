import {
  compileDocument,
  convertMillimetersToTwip,
  docxExtensions,
  effectiveRunProps,
  generateDOCX,
  generateHTML,
  generateMarkdown,
  normalizeDocument,
  parseDOCX,
  parseHTML,
  parseMarkdown,
  sectionPageSizeDefaults,
  type JSONContent,
  type SectionPropertiesOptions,
  type StylesOptions,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { projectDocumentOptions, type ProjectedFlowBox } from "@docen/docx/layout";
import { browserFontMetrics, layoutFlow, TextMeasurer, type FlowPage } from "@docen/layout";
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
import { createDefaultAddin } from "./addin";
import { mountEditBridge, type EditBridge } from "./canvas/edit-bridge";
// Side-effect: register the document-specific UI components moved out of the
// shared ui/ barrel — <docen-format-pane> (properties fallback) and
// <docen-outline> (navigation Headings tab).
import "./components/format-pane";
import "./components/outline";
import { CanvasStage } from "./canvas/stage";
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
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <docen-options-dialog part="options"></docen-options-dialog>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <input type="file" id="file-input" accept=".docx,.md,.markdown,.html,.htm" hidden />
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
  /** Cached unwrapped JSON (host.getJSON result). Invalidated on every user/doc
   *  change; recomputed lazily. Saves the editor.getJSON walk on every
   *  save/autosave/getJSON call. */
  #cachedJSON?: JSONContent;
  #jsonDirty = true;

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

  /** Paste from the system clipboard as plain text. navigator.clipboard is the
   *  reliable path; execCommand("paste") is the fallback (often blocked). */
  async #paste(): Promise<void> {
    const editor = this.editor;
    if (!editor) return;
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return;
    }
    if (text) {
      this.#bridge?.focus();
      editor.commands.insertContent(text);
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
    const baseDoc = contentAttr ? parseHTML(contentAttr) : ({} as JSONContent);
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
    this.editor?.on("transaction", this.#onTransaction);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
    this.dispatchEvent(new CustomEvent("docen:ready", { bubbles: true, composed: true }));
  }

  /** The canvas pipeline — the single render entry the bridge's transactions
   *  and the loaders share: compile → project → layout → paint, then re-arm
   *  the caret map against the fresh geometry. */
  #renderDoc(doc: JSONContent): void {
    if (!this.#stageHost) return;
    const { blocks, flow, furniture, background } = projectDocumentOptions(compileDocument(doc));
    const pages = layoutFlow(blocks, flow, this.#measurer);
    this.#pages = pages;
    this.#flow = flow;
    this.#stage ??= new CanvasStage(this.#stageHost, {
      metrics: browserFontMetrics,
      flow,
      furniture,
    });
    // A `zoom` attribute parsed before the stage existed only recorded the
    // level here — push it in before the first sync sizes the slots.
    if (this.#stage.zoom !== this.#zoom) this.#stage.setZoom(this.#zoom);
    this.#stage.sync(pages, flow, furniture, background);
    this.#bridge?.updatePages(pages, flow);
    this.#updateStatus();
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
                <fluent-menu-item data-event="save-as-html">${t("header.save-as-html", this)}</fluent-menu-item>
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

  /** Deep-merge a sectionProperties patch into the CURRENT section's sectPr and
   *  dispatch it — Word's "this section" semantics. The current section is the
   *  one holding the caret: its sectPr rides on its last paragraph (the first
   *  section-carrying paragraph at/after the caret), or, when the caret is in the
   *  final section, on doc.attrs.sectionProperties (the body-level sectPr).
   *  The dispatched transaction re-renders every page of the canvas. */
  #updateSectionGeometry(patch: SectionPropertiesOptions): void {
    const editor = this.editor;
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
    // Section count: the projection flows a single section for now; a
    // multi-section document reports its body section.
    const section = 1;
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
    // Clipboard — the selection is canvas-rendered (no DOM editor selection),
    // so copy/cut read the doc range and write the clipboard directly.
    if (name === "copy" || name === "cut") {
      void this.#copySelection(name === "cut");
      return;
    }
    if (name === "paste") {
      void this.#paste();
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
    const baseName = (this.getAttribute("filename")?.trim() || t("header.doc-name", this)).replace(
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
    const title = (this.getAttribute("filename")?.trim() || t("header.doc-name", this)).replace(
      /\.[^.]+$/,
      "",
    );
    return `<!DOCTYPE html><html lang="${escapeHtml(document.documentElement.lang || "en")}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(title)}</title></head><body>${body}</body></html>`;
  }

  /** Print the document: window.print(). */
  #print(): void {
    window.print();
  }

  /** Common load path for openDOCX/openMarkdown/openHTML: adopt a filename,
   *  replace the whole doc node. The #loadDoc wake-up transaction re-renders
   *  the canvas through the bridge. */
  #applyOpenedJSON(json: JSONContent, filename?: string): void {
    if (filename) this.setAttribute("filename", filename);
    this.#loadDoc(json);
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

  /** Serialize the current document to a DOCX buffer. */
  async saveDOCX(): Promise<Uint8Array> {
    const buffer = await generateDOCX(this.getJSON());
    return buffer as unknown as Uint8Array;
  }

  /** Serialize the current document to a Markdown string. */
  saveMarkdown(): string {
    return generateMarkdown(this.getJSON());
  }

  /** Serialize the current document to an HTML body fragment (no
   *  <html>/<!DOCTYPE> wrapper — #saveAs wraps it for a standalone file). */
  saveHTML(): string {
    return generateHTML(this.getJSON());
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
