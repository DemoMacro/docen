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
  DOCEN_CLIP_MIME,
  type JSONContent,
  type BorderOptions,
  type PageBordersOptions,
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
import { attr, customElement } from "@microsoft/fast-element";
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
  type RibbonMenuItem,
} from "../ui";
import type { ColumnsValues } from "../ui/components/workspace/columns-dialog";
import type { FontDialogPatch } from "../ui/components/workspace/font-dialog";
import { proofingLanguageName } from "../ui/components/workspace/language-dialog";
import type { LinkValues } from "../ui/components/workspace/link-dialog";
import type { PageSetupValues } from "../ui/components/workspace/page-setup-dialog";
import { createDefaultAddin, textCounter } from "./addin";
// Side-effect: register the document-specific UI components moved out of the
// shared ui/ barrel — <docen-format-pane> (properties fallback) and
// <docen-outline> (navigation Headings tab).
import "./components/format-pane";
import "./components/outline";
import {
  mountEditBridge,
  type EditBridge,
  type StoryKind,
  type StorySlot,
} from "./canvas/edit-bridge";
import { deepEq, dirtyPagesOf } from "./canvas/page-eq";
import {
  CanvasStage,
  type CanvasStageSection,
  type LaidFurnitureSection,
  layFurnitureSections,
} from "./canvas/stage";
import { documentStyles, documentTemplate, escapeHtml } from "./chrome";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import type { OutlineItem } from "./components/outline";
import { tableAncestry, WIRED_DISPATCH, type ParagraphDialogPatch } from "./extensions/commands";
import type { TablePropertiesPatch } from "./extensions/commands";
import type { BorderSideState, BordersDialogPatch } from "./extensions/commands";
import { LOCAL_HANDLED, READONLY_LIVE, SAVE_FORMATS, detectOpenFormat } from "./file-formats";
import { MARGINS, PAPER_SIZES, marginTwipsFromCss, mergeSectionProperties } from "./page-setup";
import {
  buildContextualTab,
  DEFAULT_RIBBON_TAB,
  formatMeasureTwip,
  headerFooterContextTab,
  renderRibbonFromSchema,
  ribbonActions,
  ribbonTabs,
  tableContextTabs,
  useCmUnits,
} from "./ribbon";
import {
  addSpellWord,
  checkSpelling,
  ignoreSpellWord,
  spellSuggestions,
  type SpellingIssue,
} from "./spelling";
import {
  type WatermarkPictureSpec,
  type WatermarkTextSpec,
  WATERMARK_PRESETS,
  customTextWatermarkPara,
  isWatermarkNode,
  pictureWatermarkPara,
  probeImageSize,
  stampHeaderSlots,
  watermarkPara,
} from "./watermark";

/** The word at `pos` — the non-whitespace run in the caret's text node, capped
 *  at 32 chars around the caret (CJK text has no spaces, so an uncapped run
 *  swallows the whole paragraph). Null when the caret touches no text. */
function wordRangeAt(
  doc: import("@tiptap/pm/model").Node,
  pos: number,
): { from: number; to: number } | null {
  const $pos = doc.resolve(pos);
  const node = $pos.nodeBefore?.isText
    ? $pos.nodeBefore
    : $pos.nodeAfter?.isText
      ? $pos.nodeAfter
      : null;
  if (!node) return null;
  const start = pos - ($pos.nodeBefore?.isText ? $pos.nodeBefore.nodeSize : 0);
  const chars = node.text ?? "";
  const at = pos - start;
  let from = at;
  while (from > 0 && !/\s/.test(chars.charAt(from - 1))) from--;
  let to = at;
  while (to < chars.length && !/\s/.test(chars.charAt(to))) to++;
  if (to - from > 32) {
    from = Math.max(from, at - 16);
    to = Math.min(to, at + 16);
  }
  return from < to ? { from: start + from, to: start + to } : null;
}

/** Split buttons whose face carries no command of its own — the handler only
 *  exists for the drop-down variants' values (Word's menu buttons; a face
 *  click opens the menu instead of emitting a valueless command). */
const FACE_ONLY_SPLITS: ReadonlySet<string> = new Set(["autofit", "columns"]);

/** Word's Match Destination Formatting in paste-options form: the block
 *  structure survives (lists, tables, images), the source's run formatting
 *  (bold/italic/color/…) is dropped so the destination's takes over. */
function stripRunMarks(nodes: JSONContent[]): JSONContent[] {
  return nodes.map((n) => {
    const content = n.content ? stripRunMarks(n.content) : undefined;
    return n.type === "text" ? { ...n, marks: undefined } : content ? { ...n, content } : n;
  });
}

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
/**
 * Task pane identifiers, mirroring the Office `<TaskpaneId>` concept. The host
 * ships two built-in panes: `navigation` (start/left) and `properties` (end/right).
 */
export type TaskPaneId = "navigation" | "properties" | "comments" | "clipboard" | "proofing";

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
  /** Leafer engine debug overlay — "bounds" | "hit" | "repaint" | "on". */
  @attr debug?: string;
  /** The document view (Word's View tab): "print" | "web" | "draft" | "read".
   *  Anything else falls back to "print". */
  @attr view?: string;

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

  viewChanged(): void {
    this.#applyView();
  }

  addinsAttrChanged(): void {
    this.#applyAddinsAttr();
  }

  themeChanged(): void {
    this.#applyThemeAttr(this.theme ?? "");
  }

  debugChanged(): void {
    this.#stage?.setDebug(this.debug);
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

  /** Ctrl+wheel zoom over the page area (Word/Office behavior). Captured ahead
   *  of the stage shell's wheel handling, which stops propagation for its own
   *  scroll — plain wheel keeps scrolling; only the Ctrl chord zooms. */
  readonly #onWheel = (event: WheelEvent): void => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    this.#setZoom(this.#zoom + (event.deltaY < 0 ? 10 : -10));
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
    // F12 = Save As (Word).
    if (event.key === "F12") {
      event.preventDefault();
      void this.#saveAs();
      return;
    }
    // F7 = Spelling & Grammar (Word).
    if (event.key === "F7") {
      event.preventDefault();
      this.#onCommand(new CustomEvent("command", { detail: { event: "spell-check" } }));
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
    // Ctrl+S saves (Word) — before the input gate, a save applies everywhere.
    if (event.key === "s" || event.key === "S") {
      event.preventDefault();
      if (!this.#emitCancelable("docen:save")) void this.#saveAs();
      return;
    }
    // Ctrl+P prints the canvas pages (Word) — not the browser's DOM print.
    if (event.key === "p" || event.key === "P") {
      event.preventDefault();
      if (!this.#emitCancelable("docen:print")) this.#print();
      return;
    }
    // composedPath()[0] is the real target inside the shadow DOM (e.g. a combobox input).
    const target = event.composedPath()[0] as HTMLElement | null;
    if (
      target instanceof HTMLElement &&
      target.closest("input, textarea, docen-ribbon-combobox") &&
      // The bridge textarea IS the document input — zoom must stay live with
      // the caret in the document.
      !target.closest("[data-docen-bridge-input]")
    )
      return;
    // Ctrl+K opens the Link dialog (Word) — after the input gate, so a
    // combobox keystroke is never hijacked.
    if (event.key === "k" || event.key === "K") {
      event.preventDefault();
      this.#insertLink();
      return;
    }
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
    this.#bridge?.scrollIntoView(editor.state.selection.from);
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
    if (action === "find-next" || action === "replace-next") {
      this.#bridge?.scrollIntoView(editor.state.selection.from);
    }
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

  /** Paste from the system clipboard. The docen lane wins (a copy from a
   *  docen editor round-trips losslessly through the custom MIME — Chrome
   *  reads it back as a web custom format), then text/html — styled paste
   *  through the schema's parse rules — then plain text; `textOnly` (the
   *  menu's Keep Text Only) skips the rich legs. navigator.clipboard is the
   *  reliable path; execCommand("paste") is the fallback (often blocked). */
  async #paste(textOnly = false): Promise<void> {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    this.#bridge?.focus();
    const docenType = textOnly ? null : `web ${DOCEN_CLIP_MIME}`;
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (docenType && item.types.includes(docenType)) {
          const raw = await (await item.getType(docenType)).text();
          if (raw && this.#bridge?.insertSlicePayload(raw)) {
            const plain =
              item.types.includes("text/plain") &&
              (await (await item.getType("text/plain")).text());
            if (plain) this.#showPasteOptions({ kind: "slice", raw, text: plain });
            return;
          }
        }
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
            const plain = item.types.includes("text/plain")
              ? await (await item.getType("text/plain")).text()
              : "";
            this.#showPasteOptions({ kind: "html", raw: text, text: plain });
            return;
          }
        } else {
          // Chromium never persists a copy EVENT's custom types to the system
          // clipboard (the `web ` spelling only survives the async write API),
          // so a same-page Ctrl+C → context-menu Paste round-trip can't see
          // the docen lane in read(). When the plain text still matches the
          // pinned in-editor copy, the slice payload rides memory instead; a
          // copy made elsewhere produces different text and the fallback
          // correctly stays out.
          const pinned = textOnly ? null : this.#bridge?.copiedSlice();
          if (pinned && pinned.text === text && this.#bridge?.insertSlicePayload(pinned.payload)) {
            return;
          }
          editor.commands.insertContent(text);
          return;
        }
      }
    } catch {
      // read() may be denied (permission policy) — fall through to readText.
    }
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        const pinned = textOnly ? null : this.#bridge?.copiedSlice();
        if (pinned && pinned.text === text && this.#bridge?.insertSlicePayload(pinned.payload)) {
          this.#showPasteOptions({ kind: "slice", raw: pinned.payload, text });
          return;
        }
        editor.commands.insertContent(text);
      }
    } catch {
      /* clipboard unavailable — nothing to paste */
    }
  }

  /** The clipboard content behind the live paste-options bar — the source
   *  the three picks replay in their picked form (Word's paste options). */
  #pasteBar?: HTMLElement;
  #pasteSource?: { kind: "slice" | "html"; raw: string; text: string };

  /** Word's paste-options bar — after a rich paste it hangs below the pasted
   *  content; the three picks undo the insertion and replay the same
   *  clipboard content in the picked form (source / destination-matched /
   *  text only). */
  #showPasteOptions(source: { kind: "slice" | "html"; raw: string; text: string }): void {
    this.#hidePasteOptions();
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    const anchor = editor ? this.#bridge?.pasteAnchorRect(editor.state.selection.from) : null;
    if (!anchor) return;
    this.#pasteSource = source;
    const bar = document.createElement("div");
    bar.setAttribute("data-docen-overlay", "");
    Object.assign(bar.style, {
      position: "absolute",
      zIndex: "7",
      display: "flex",
      gap: "2px",
      padding: "3px",
      background: "var(--docen-color-bg, #ffffff)",
      border: "1px solid var(--docen-color-divider, #e2e2e2)",
      borderRadius: "var(--borderRadiusMedium, 6px)",
      boxShadow: "var(--shadow4, 0 4px 8px rgba(0,0,0,.14))",
    } satisfies Partial<CSSStyleDeclaration>);
    const picks: Array<{ mode: "source" | "match" | "text"; key: string }> = [
      { mode: "source", key: "ribbon.opt.keep-source" },
      { mode: "match", key: "ribbon.opt.match-dest" },
      { mode: "text", key: "ribbon.opt.keep-text-only" },
    ];
    for (const pick of picks) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = t(pick.key, this);
      Object.assign(btn.style, {
        padding: "3px 8px",
        border: "none",
        background: "transparent",
        borderRadius: "4px",
        cursor: "pointer",
        whiteSpace: "nowrap",
        fontSize: "12px",
        lineHeight: "1.4",
        color: "inherit",
        fontFamily: "inherit",
      } satisfies Partial<CSSStyleDeclaration>);
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.#hidePasteOptions();
        this.#replayPaste(pick.mode);
      });
      bar.append(btn);
    }
    anchor.frame.append(bar);
    // Below the pasted content's last line, clamped into the frame.
    Object.assign(bar.style, {
      left: `${Math.max(0, anchor.left)}px`,
      top: `${anchor.top + anchor.height + 4}px`,
    });
    this.#pasteBar = bar;
    document.addEventListener("mousedown", this.#pasteBarDismiss, true);
    document.addEventListener("keydown", this.#pasteBarDismiss, true);
  }

  #hidePasteOptions(): void {
    document.removeEventListener("mousedown", this.#pasteBarDismiss, true);
    document.removeEventListener("keydown", this.#pasteBarDismiss, true);
    this.#pasteBar?.remove();
    this.#pasteBar = undefined;
    this.#pasteSource = undefined;
  }

  /** A click outside the bar or Escape dismisses it (Word keeps the paste
   *  itself — only the options bar goes away). Clicks inside the bar fall
   *  through to the pick buttons. */
  readonly #pasteBarDismiss = (event: Event): void => {
    if (event instanceof KeyboardEvent && event.key !== "Escape") return;
    if (event instanceof MouseEvent && this.#pasteBar?.contains(event.target as Node)) return;
    this.#hidePasteOptions();
  };

  /** The Office Clipboard's session collection (newest first, Word's 24-item
   *  cap) — fed by the bridge's onClipboardCollect, rendered by the pane. */
  #clipboardItems: { text: string; payload: string | null }[] = [];

  #collectClipboardItem(item: { text: string; payload: string | null }): void {
    // A re-copy of the newest item keeps the pane's order stable.
    if (this.#clipboardItems[0]?.text === item.text) return;
    this.#clipboardItems.unshift({ ...item });
    if (this.#clipboardItems.length > 24) this.#clipboardItems.length = 24;
    this.#syncClipboardPane();
  }

  #syncClipboardPane(): void {
    const pane = this.shadowRoot?.querySelector("docen-clipboard-pane");
    if (pane) (pane as unknown as { entries: unknown[] }).entries = [...this.#clipboardItems];
  }

  #pasteClipboardEntry(entry: { text: string; payload: string | null } | null): void {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor || !entry) return;
    this.#bridge?.focus();
    if (entry.payload && this.#bridge?.insertSlicePayload(entry.payload)) return;
    editor.commands.insertContent(entry.text);
  }

  readonly #onClipboardPaste = (event: Event): void => {
    this.#pasteClipboardEntry(
      (event as CustomEvent<{ text: string; payload: string | null }>).detail,
    );
  };

  /** Word's Paste All — items land in collection order (oldest first). */
  readonly #onClipboardPasteAll = (): void => {
    for (const entry of [...this.#clipboardItems].reverse()) this.#pasteClipboardEntry(entry);
  };

  readonly #onClipboardClear = (): void => {
    this.#clipboardItems = [];
  };

  /** A paste-options pick: undo the insertion, then replay the same clipboard
   *  content in the picked form. "match" keeps the block structure (lists,
   *  tables, links) but drops the source's run formatting so the destination's
   *  takes over; "text" inserts the plain text. */
  #replayPaste(mode: "source" | "match" | "text"): void {
    const source = this.#pasteSource;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!source || !editor) return;
    editor.commands.undo();
    if (mode === "text") {
      editor.commands.insertContent(source.text);
      return;
    }
    if (source.kind === "slice") {
      if (mode === "source" && this.#bridge?.insertSlicePayload(source.raw)) return;
      // A slice replayed destination-matched: strip its run marks.
      try {
        const parsed = JSON.parse(source.raw) as { content?: JSONContent[] };
        if (parsed.content) {
          editor.commands.insertContent(stripRunMarks(parsed.content));
        }
      } catch {
        /* unparsable payload — the undo already restored the prior state */
      }
      return;
    }
    const body = new DOMParser().parseFromString(source.raw, "text/html").body;
    const json = parseHTMLBody(body, editor.state.schema);
    const content = (json.content ?? []).filter((n) => n.type !== "text" || n.text);
    if (content.length)
      editor.commands.insertContent(mode === "match" ? stripRunMarks(content) : content);
  }

  /** Editing → Select menu. "all" uses the official selectAll() command;
   *  "objects"/"similar" are placeholders. */
  #select(value?: string): void {
    // The story the caret lives in — Ctrl+A selects the story text, the menu
    // command must agree with it (a stale main-doc range would be invisible).
    const editor = this.#bridge?.activeEditor() ?? this.editor;
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
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor || editor.state.selection.empty) return;
    // Probe one character into the selection: $from sits on the boundary,
    // and ResolvedPos.marks() reads the character BEFORE the position — the
    // first selected character's marks (e.g. bold stamped on [from,to))
    // would be lost.
    this.#painterMarks = editor.state.doc.resolve(editor.state.selection.from + 1).marks();
    this.toggleAttribute("format-painter", true);
    const onUp = (): void => {
      const ed = this.#bridge?.activeEditor() ?? this.editor;
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
      this.#syncContextTabs();
      // After #syncContextTabs: the first transaction that enters a table is
      // also the one that appends the Table Layout panel — the combos only
      // exist from that pass on.
      this.#syncCellSize();
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
    // The proofing surfaces: the status-bar book opens the pane; the pane's
    // actions (replace/ignore/add/step) come back as events.
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.addEventListener("spellcheck:open", () =>
        this.#onCommand(new CustomEvent("command", { detail: { event: "spell-check" } })),
      );
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:replace", ((event: CustomEvent<string>) =>
        this.#replaceSpellingIssue(event.detail)) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:ignore-all", () => this.#ignoreSpelling("ignore"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:add", () => this.#ignoreSpelling("add"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:nav", ((event: CustomEvent<number>) =>
        this.#gotoSpellingIssue(this.#spellingActive + event.detail)) as EventListener);

    this.#stageHost = this.shadowRoot!.querySelector<HTMLElement>(".docen-canvas") ?? undefined;
    this.#stageHost?.addEventListener("wheel", this.#onWheel as EventListener, {
      capture: true,
      passive: false,
    });
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
      // The textarea must live outside docen-context-menu (fluent-menu eats
      // Space/Enter) — the input layer at the shadow root is menu-free.
      inputHost: this.shadowRoot!.querySelector<HTMLElement>(".input-layer")!,
      content: initialDoc,
      onDoc: (json) => this.#renderDoc(json),
      pageHost: (page) => this.#stage?.slotAt(page)?.parentElement ?? null,
      extensions: [...docxExtensions, ...(defaultAddin.extensions ?? [])],
      scale: () => this.#stage?.scale() ?? 1,
      // Word's paste-options bar hangs after every rich paste; the clipboard
      // pane collects each in-editor copy/cut.
      onRichPaste: (source) => this.#showPasteOptions(source),
      onClipboardCollect: (item) => this.#collectClipboardItem(item),
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
          this.#showHeaderFooterContextTab();
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
      // `#name` links are bookmark anchors — the host owns the in-page jump.
      onInternalAnchor: (name) => this.#jumpToBookmark(name),
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
    // Right-click → Word's context menu. Captured on the shadow root so the
    // items are built before <docen-context-menu>'s own capture handler opens
    // the Fluent menu (capture runs outermost-first).
    this.shadowRoot!.addEventListener("contextmenu", this.#onContextMenu as EventListener, true);
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
    // Office Clipboard pane → paste one entry / paste all / clear.
    this.addEventListener("clipboard:paste", this.#onClipboardPaste as EventListener);
    this.addEventListener("clipboard:paste-all", this.#onClipboardPasteAll as EventListener);
    this.addEventListener("clipboard:clear", this.#onClipboardClear as EventListener);
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
    // Options dialog — ok (UI language + theme).
    this.shadowRoot!.querySelector("docen-options-dialog")?.addEventListener(
      "options:ok",
      this.#onOptionsOk as EventListener,
    );
    // Language dialog — commit the selection's proofing language (w:lang).
    this.shadowRoot!.querySelector("docen-language-dialog")?.addEventListener(
      "language:ok",
      this.#onLanguageOk as EventListener,
    );
    // Phonetic guide dialog — split the selection into per-character ruby
    // runs, or strip the guides off it.
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:ok",
      this.#onPhoneticOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:clear",
      this.#onPhoneticClear as EventListener,
    );
    // Status-bar language item — open the language dialog (Word semantics).
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "language:open",
      this.#onLanguageOpen as EventListener,
    );
    // Symbol dialog — insert the picked character at the caret.
    this.shadowRoot!.querySelector("docen-symbol-dialog")?.addEventListener(
      "symbol:insert",
      this.#onSymbolInsert as EventListener,
    );
    // Paragraph dialog — stamp the committed patch onto the selection.
    this.shadowRoot!.querySelector("docen-paragraph-dialog")?.addEventListener(
      "paragraph:ok",
      this.#onParagraphOk as EventListener,
    );
    // Page Setup dialog — write the committed geometry into the current
    // section (the Custom Margins / More Paper Sizes entries open it).
    this.shadowRoot!.querySelector("docen-page-setup-dialog")?.addEventListener(
      "page-setup:ok",
      this.#onPageSetupOk as EventListener,
    );
    // Table grid — insert the picked shape through insert-table.
    this.shadowRoot!.querySelector("docen-table-dialog")?.addEventListener(
      "table-grid:insert",
      this.#onTableInsert as EventListener,
    );
    // Columns dialog — write the committed layout into the current section.
    this.shadowRoot!.querySelector("docen-columns-dialog")?.addEventListener(
      "columns:ok",
      this.#onColumnsOk as EventListener,
    );
    // Link dialog — commit the hyperlink (mark / replace / insert / remove).
    this.shadowRoot!.querySelector("docen-link-dialog")?.addEventListener(
      "link:ok",
      this.#onLinkOk as EventListener,
    );
    // Zoom dialog — apply the preset or free percent; the status-bar percent
    // click opens it.
    this.shadowRoot!.querySelector("docen-zoom-dialog")?.addEventListener(
      "zoom:ok",
      this.#onZoomOk as EventListener,
    );
    // Paste Special dialog — the format pick re-runs #paste in that mode.
    this.shadowRoot!.querySelector("docen-paste-special-dialog")?.addEventListener(
      "paste-special:ok",
      this.#onPasteSpecialOk as EventListener,
    );
    // Font dialog — stamp the committed run state onto the selection.
    this.shadowRoot!.querySelector("docen-font-dialog")?.addEventListener(
      "font:ok",
      this.#onFontDialogOk as EventListener,
    );
    // Table Properties dialog — rewrite the caret table's alignment/indent.
    this.shadowRoot!.querySelector("docen-table-properties-dialog")?.addEventListener(
      "table-properties:ok",
      this.#onTablePropertiesOk as EventListener,
    );
    // Borders and Shading dialog — stamp the border/page/shading tab.
    this.shadowRoot!.querySelector("docen-borders-shading-dialog")?.addEventListener(
      "borders-shading:ok",
      this.#onBordersShadingOk as EventListener,
    );
    // Custom watermark dialog — clear/stamp the header watermark.
    this.shadowRoot!.querySelector("docen-watermark-dialog")?.addEventListener(
      "watermark:ok",
      this.#onWatermarkOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "zoom:open",
      this.#onZoomOpen as EventListener,
    );
    // Status-bar view shortcuts (Word's Reading / Print Layout / Web Layout
    // buttons) — the same `view` attribute the ribbon's View tab writes.
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "view:select",
      this.#onViewSelect as EventListener,
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
    // Selection moves repaint the anchored comment card (Word highlights the
    // card whose range the caret sits in).
    this.editor?.on("selectionUpdate", this.#syncActiveCommentCard);
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
    this.#hideHeaderFooterContextTab();
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

  /** The active view, normalized (an unknown attr value reads as print). */
  #viewMode(): "print" | "web" | "draft" | "read" {
    return this.view === "web" || this.view === "draft" || this.view === "read"
      ? this.view
      : "print";
  }

  /** Apply the `view` attribute: restage the render mode, re-project (the
   *  continuous views re-layout at the viewport width; Draft drops the
   *  furniture insets), trim the chrome in Read Mode, and sync the status
   *  bar. Word's Read Mode is read-only; the other views keep editability. */
  #applyView(): void {
    const mode = this.#viewMode();
    this.#stage?.setViewMode(mode);
    this.#syncReadChrome(mode === "read");
    if (this.editor) {
      const editable = this.editable !== "false" && mode !== "read";
      if (this.editor.isEditable !== editable) {
        this.editor.setEditable(editable);
        this.#syncEditModeMenu();
      }
    }
    this.#renderDoc(this.getJSON());
    this.#updateStatus();
  }

  /** Read Mode trims the chrome to the document (Word hides the ribbon and
   *  most of the tab row). */
  #syncReadChrome(read: boolean): void {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (ribbon) ribbon.toggleAttribute("hidden", read);
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
    // The continuous views (Web Layout / Read Mode) re-box every section to
    // the viewport width and lay it as ONE unbounded page — Word's web view
    // has no page breaks, and its text width follows the window (page margins
    // kept as the gutters). Furniture is a print concept: no insets. Columns
    // stay a print-layout feature in this pass.
    const mode = this.#viewMode();
    const continuous = mode === "web" || mode === "read";
    if (continuous) {
      // The scroll surface's width (the document area — the stage host is
      // width:fit-content and only reports the pages' own width) minus the
      // page gutter on each side.
      const area = this.shadowRoot?.querySelector<HTMLElement>("docen-document-area");
      const availW = Math.max(320, (area?.clientWidth ?? 794) - 48);
      for (const section of stageSections) {
        // Columns stay a print-layout feature: drop them at the source so
        // every downstream consumer (flow opts, separator painting) lays a
        // single-column stream.
        section.columns = undefined;
        const marginL = section.flow.contentLeftPx;
        const marginR =
          section.flow.pageWidthPx - section.flow.contentLeftPx - section.flow.contentWidthPx;
        section.flow = {
          ...section.flow,
          pageWidthPx: availW,
          contentWidthPx: Math.max(200, availW - marginL - marginR),
        };
      }
    }
    const laidFurniture = layFurnitureSections(stageSections, browserFontMetrics);
    stageSections.forEach((section, i) => {
      section.furnitureLaid = laidFurniture[i];
    });
    const flowSections = stageSections.map((section) => {
      const pageInsets = continuous
        ? undefined
        : this.#pageInsets(section.flow, section.furniture, section.furnitureLaid);
      return {
        blocks: section.blocks,
        opts: {
          ...section.flow,
          columns: section.columns,
          footnoteDefinitions: section.footnoteDefinitions,
          endnoteDefinitions: section.endnoteDefinitions,
          ...(continuous ? { unbounded: true, contentHeightPx: 1_000_000 } : {}),
          ...(pageInsets ? { pageInsets } : {}),
        },
      };
    });
    const { pages, sectionOfPage } = layoutFlowSections(flowSections, this.#measurer);
    if (continuous) {
      // Size each continuous page to where its content actually ends (the
      // unbounded layout reports the content bottom) plus the bottom margin.
      pages.forEach((page, i) => {
        const section = stageSections[sectionOfPage[i] ?? 0];
        if (!section || page.contentBottomPx == null) return;
        const flow = section.flow;
        const bottomMargin = flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx;
        flow.pageHeightPx = Math.max(
          flow.pageHeightPx,
          Math.ceil(page.contentBottomPx) + flow.contentTopPx + bottomMargin,
        );
      });
    }
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
   *  the caret map against the fresh geometry. Page-level diff: only pages
   *  whose laid-out content changed repaint (the rest keep their canvas), so
   *  a keystroke costs one page, not one per scrolled-into-view page. */
  #renderDoc(doc: JSONContent): void {
    if (!this.#stageHost) return;
    const prev = this.#lastRun;
    const run = { ...this.#projectAndLayout(doc), viewMode: this.#viewMode() };
    this.#lastRun = run;
    this.#pages = run.pages;
    this.#sectionOfPage = run.sectionOfPage;
    this.#flow = run.sections[0]?.flow;
    this.#stage ??= new CanvasStage(this.#stageHost, {
      metrics: browserFontMetrics,
      sections: run.sections,
      sectionOfPage: run.sectionOfPage,
      background: run.background,
    });
    // A debug attribute stamped before the first render lands here.
    if (this.debug) this.#stage.setDebug(this.debug);
    this.#stage.setMarksLabels({
      pageBreak: t("marks.pageBreak", this),
      sectionBreak: t("marks.sectionBreak", this),
    });
    // A `zoom` attribute parsed before the stage existed only recorded the
    // level here — push it in before the first sync sizes the slots. The
    // `show-marks` and `view` attributes get the same once-over (idempotent
    // setters; the read-only + chrome trimming rides #applyView's gate).
    if (this.#stage.zoom !== this.#zoom) this.#stage.setZoom(this.#zoom);
    if (this.hasAttribute("show-marks")) this.#stage.setShowMarks(true);
    if (this.#stage.viewMode !== this.#viewMode()) {
      this.#stage.setViewMode(this.#viewMode());
      this.#syncReadChrome(this.#viewMode() === "read");
      if (this.editor) {
        const editable = this.editable !== "false" && this.#viewMode() !== "read";
        this.editor.setEditable(editable);
        this.#syncEditModeMenu();
      }
    }
    // Anything structural (section geometry, page background, section count)
    // repaints everything — the per-page diff only skips pages whose
    // placement, section, AND background are all unchanged. A page-count
    // shift is deliberately NOT structural: deleting across a pagination
    // boundary bounces the count between renders, and a full repaint per
    // bounce is the visible flicker of holding Backspace. dirtyPagesOf
    // covers count changes positionally (pages past either end stay dirty;
    // the trailing slots' lifecycles are handled in sync). The overlapping
    // page range still compares its section map: a deletion may pull a later
    // section onto an existing page slot, which changes its flow/furniture.
    // Furniture compares on the projected options, not the laid stacks (the
    // stacks derive from them plus the already-compared flow width); the
    // frame CSS (background/borders) re-stamps on every sync and needs no
    // diff.
    const structural =
      !prev ||
      prev.viewMode !== run.viewMode ||
      prev.sectionOfPage.some(
        (section, index) =>
          index < run.sectionOfPage.length && section !== run.sectionOfPage[index],
      ) ||
      prev.background?.color !== run.background?.color ||
      prev.sections.length !== run.sections.length ||
      prev.sections.some((s, i) => !deepEq(s.flow, run.sections[i]!.flow)) ||
      prev.sections.some((s, i) => !deepEq(s.furniture, run.sections[i]!.furniture)) ||
      prev.sections.some((s, i) => !deepEq(s.lineNumbers, run.sections[i]!.lineNumbers)) ||
      prev.sections.some((s, i) => !deepEq(s.columns, run.sections[i]!.columns));
    const dirty = structural ? undefined : dirtyPagesOf(prev.pages, run.pages);
    this.#stage.sync(run.pages, run.sections, run.sectionOfPage, run.background, dirty);
    this.#bridge?.updatePages(run.pages, this.#pageOriginOf(run.sections, run.sectionOfPage));
    this.#updateStatus();
    this.#syncCommentsPane();
    this.#scheduleSpellCheck();
    this.#syncStatusLanguage();
  }

  /** The previous render's flow result — the diff base for the next one. */
  #lastRun?: {
    pages: FlowPage[];
    sectionOfPage: number[];
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
    viewMode: "print" | "web" | "draft" | "read";
  };

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
      ?.removeEventListener("language:open", this.#onLanguageOpen as EventListener);
    this.shadowRoot
      ?.querySelector("docen-language-dialog")
      ?.removeEventListener("language:ok", this.#onLanguageOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:ok", this.#onPhoneticOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:clear", this.#onPhoneticClear as EventListener);
    this.shadowRoot
      ?.querySelector("docen-symbol-dialog")
      ?.removeEventListener("symbol:insert", this.#onSymbolInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paragraph-dialog")
      ?.removeEventListener("paragraph:ok", this.#onParagraphOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paste-special-dialog")
      ?.removeEventListener("paste-special:ok", this.#onPasteSpecialOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-font-dialog")
      ?.removeEventListener("font:ok", this.#onFontDialogOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-properties-dialog")
      ?.removeEventListener("table-properties:ok", this.#onTablePropertiesOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-borders-shading-dialog")
      ?.removeEventListener("borders-shading:ok", this.#onBordersShadingOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-watermark-dialog")
      ?.removeEventListener("watermark:ok", this.#onWatermarkOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-page-setup-dialog")
      ?.removeEventListener("page-setup:ok", this.#onPageSetupOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-dialog")
      ?.removeEventListener("table-grid:insert", this.#onTableInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-columns-dialog")
      ?.removeEventListener("columns:ok", this.#onColumnsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-link-dialog")
      ?.removeEventListener("link:ok", this.#onLinkOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-zoom-dialog")
      ?.removeEventListener("zoom:ok", this.#onZoomOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("zoom:open", this.#onZoomOpen as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.removeEventListener("zoom:change", this.#onZoomChange as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("view:select", this.#onViewSelect as EventListener);
    this.#stageHost?.removeEventListener("wheel", this.#onWheel as EventListener, {
      capture: true,
    });
    this.editor?.off("transaction", this.#onTransaction);
    this.editor?.off("selectionUpdate", this.#syncActiveCommentCard);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.shadowRoot
      ?.querySelector("docen-ribbon")
      ?.removeEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    this.#fontSyncCleanup?.();
    this.#fontSyncCleanup = undefined;
    this.#stopFormatPainter();
    clearTimeout(this.#searchTimer);
    if (this.#spellingTimer != null) clearTimeout(this.#spellingTimer);
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
    // The ribbon DOM was rebuilt from scratch — drop the stale context-tab
    // tracking, then re-append them if the selection is inside a table.
    this.#contextTabIds.clear();
    this.#syncContextTabs();
    this.#syncCellSize();
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
        "docen-ribbon-button[event], docen-ribbon-split-button[event], docen-ribbon-toggle-button[event], docen-ribbon-menu[event]",
      )
      .forEach((el) => {
        const event = el.getAttribute("event");
        if (!event) return;
        // A composite (split/menu) stays live while ANY drop-down variant
        // resolves to a wired command — greying the host would bury its live
        // items (the AutoFit split's face has no action, its three variants
        // do). A face-only split keeps its caret and opens the menu instead.
        const liveItems =
          el.tagName === "DOCEN-RIBBON-SPLIT-BUTTON" || el.tagName === "DOCEN-RIBBON-MENU"
            ? this.#ribbonMenuItems(el).some(
                (item) => !item.disabled && wired.has(item.event ?? event),
              )
            : false;
        if (wired.has(event) || liveItems) {
          el.removeAttribute("disabled");
          // A face with no action of its own opens the drop-down instead:
          // either its own event is unwired while the variants are live, or
          // the handler only exists for the variants' values (Columns).
          const faceOnly =
            el.tagName === "DOCEN-RIBBON-SPLIT-BUTTON" && FACE_ONLY_SPLITS.has(event);
          if (faceOnly || (liveItems && !wired.has(event)))
            el.setAttribute("primary-opens-menu", "");
          else el.removeAttribute("primary-opens-menu");
        } else {
          el.setAttribute("disabled", "");
          el.removeAttribute("primary-opens-menu");
        }
      });
  }

  /** A composite control's parsed `items` attribute (menu variants), empty on
   *  malformed JSON so a typo greys the control rather than crashing. */
  #ribbonMenuItems(el: HTMLElement): { disabled?: boolean; event?: string }[] {
    try {
      return JSON.parse(el.getAttribute("items") ?? "[]") as {
        disabled?: boolean;
        event?: string;
      }[];
    } catch {
      return [];
    }
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
    // While a story is open the chrome re-stamp may have rebuilt the ribbon —
    // re-hang the context tab and mirror the flags into its checkboxes.
    if (this.#storyKind != null) {
      this.#showHeaderFooterContextTab();
      const titleCb = this.shadowRoot?.querySelector(
        'docen-ribbon-checkbox[event="header-option"][value="title-page"]',
      );
      titleCb?.toggleAttribute("checked", !!sp?.titlePage);
      const oddEvenCb = this.shadowRoot?.querySelector(
        'docen-ribbon-checkbox[event="header-option"][value="odd-even"]',
      );
      oddEvenCb?.toggleAttribute("checked", !!sp?.evenAndOddHeaders);
    }
  }

  /** Word's Header & Footer Tools — append the contextual tab while a story
   *  is open and activate it (Word drops you on the tab); idempotent across
   *  chrome re-stamps. */
  #showHeaderFooterContextTab(): void {
    const root = this.shadowRoot;
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    if (tablist.querySelector("#header-footer-tab")) return;
    const scope = root.querySelector("docen-workspace") ?? this;
    const built = buildContextualTab(headerFooterContextTab(), scope);
    tablist.append(built.tab);
    ribbon.append(built.panel);
    tablist.setAttribute("activeid", "header-footer-tab");
    this.#applyRibbonGreying();
  }

  #hideHeaderFooterContextTab(): void {
    const root = this.shadowRoot;
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    if (!tablist.querySelector("#header-footer-tab")) return;
    if (tablist.getAttribute("activeid") === "header-footer-tab")
      tablist.setAttribute("activeid", DEFAULT_RIBBON_TAB);
    tablist.querySelector("#header-footer-tab")?.remove();
    ribbon.querySelector('docen-ribbon-panel[value="header-footer-tab"]')?.remove();
    this.#applyRibbonGreying();
  }

  /** Mirror the caret cell's live width/height into the Cell Size combos —
   *  Word behavior: the boxes report the selection's column width and row
   *  height (in the locale's unit system), not a fixed default. Runs on every
   *  chrome re-stamp and transaction (via #setupFontSync). */
  #syncCellSize(): void {
    const root = this.shadowRoot;
    const widthEl = root?.querySelector('docen-ribbon-combobox[event="cell-width"]');
    const heightEl = root?.querySelector('docen-ribbon-combobox[event="cell-height"]');
    const editor = this.editor;
    if ((!widthEl && !heightEl) || !editor) return;
    const anchor = tableAncestry(editor.state);
    if (!anchor) return;
    const scope = root?.querySelector("docen-workspace") ?? this;
    const { $from } = editor.state.selection;
    if (widthEl) {
      const widths = ($from.node(anchor.tableAt).attrs as { columnWidths?: number[] | null })
        .columnWidths;
      const col = $from.index(anchor.rowAt);
      const tw = widths != null && col < widths.length ? widths[col] : undefined;
      widthEl.setAttribute("value", tw != null ? formatMeasureTwip(tw, scope) : "");
    }
    if (heightEl) {
      const h = ($from.node(anchor.rowAt).attrs as { height?: { value?: number } | null }).height;
      heightEl.setAttribute(
        "value",
        h?.value != null ? formatMeasureTwip(h.value, scope) : useCmUnits(scope) ? "自动" : "auto",
      );
    }
  }

  /** Contextual tab ids currently appended to the ribbon (Word's Table Tools).
   *  Non-empty ⇔ the selection is inside a table; #syncContextTabs diffs this
   *  against that fact so the per-transaction pass is a cheap equality check. */
  #contextTabIds = new Set<string>();

  /** Word's Table Tools — append/remove the contextual Table Design/Layout tabs
   *  as the selection enters/leaves a table. Runs per transaction (via
   *  #setupFontSync) and after every chrome re-stamp (#renderChrome, which
   *  clears the tracking set because the ribbon DOM was rebuilt). */
  #syncContextTabs(): void {
    const root = this.shadowRoot;
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    const inside = this.editor ? tableAncestry(this.editor.state) !== null : false;
    const present = this.#contextTabIds;
    if (inside === present.size > 0) return;
    const scope = root.querySelector("docen-workspace") ?? this;
    if (inside) {
      for (const tab of tableContextTabs(scope)) {
        const built = buildContextualTab(tab, scope);
        tablist.append(built.tab);
        ribbon.append(built.panel);
        present.add(tab.id);
      }
      // Word activates the context tabs as the caret enters their context.
      tablist.setAttribute("activeid", "table-design");
    } else {
      // Fall back off the context tabs BEFORE removing them so the tablist
      // never holds an activeid with no matching tab.
      if (present.has(tablist.getAttribute("activeid") ?? "")) {
        tablist.setAttribute("activeid", DEFAULT_RIBBON_TAB);
      }
      for (const id of present) {
        tablist.querySelector(`#${id}`)?.remove();
        ribbon.querySelector(`docen-ribbon-panel[value="${id}"]`)?.remove();
      }
      present.clear();
    }
    this.#applyRibbonGreying();
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
    // The status-bar language mirrors the caret's proofing language (Word).
    if (props.transaction.selectionSet) this.#syncStatusLanguage();
    if (props.transaction.docChanged) {
      this.#jsonDirty = true;
      this.dispatchEvent(
        new CustomEvent("docen:change", { bubbles: true, composed: true, detail: { dirty: true } }),
      );
    }
  };

  // ---- Spelling (Review → Spelling & Grammar) ----
  // The host owns the check: debounced after every render, its issue list
  // feeds the squiggle overlay, the proofing pane, and the status-bar book.

  #spellingIssues: SpellingIssue[] = [];
  /** The pane's active issue (document order); -1 = nothing selected. */
  #spellingActive = -1;
  #spellingTimer: ReturnType<typeof setTimeout> | null = null;

  #scheduleSpellCheck(): void {
    if (this.#spellingTimer != null) clearTimeout(this.#spellingTimer);
    this.#spellingTimer = setTimeout(() => {
      this.#spellingTimer = null;
      this.#runSpellCheck();
    }, 400);
  }

  #runSpellCheck(): void {
    const editor = this.editor;
    if (!editor) return;
    this.#spellingIssues = checkSpelling(editor.state.doc);
    this.#bridge?.setSpellingIssues(this.#spellingIssues);
    this.#spellingActive = this.#spellingIssues.length ? 0 : -1;
    const bar = this.shadowRoot?.querySelector("docen-status-bar");
    bar?.setAttribute("proofing", this.#spellingIssues.length ? "issues" : "ok");
    this.#syncSpellingPane();
  }

  /** Push the active issue (with its suggestions, computed here — the pane
   *  stays data-only) into the proofing pane when it's open. */
  #syncSpellingPane(): void {
    const pane = this.shadowRoot?.querySelector("docen-spelling-pane") as
      | (HTMLElement & {
          entries: Array<{ word: string; suggestions: string[] }>;
          active: number;
          total: number;
        })
      | null;
    if (!pane) return;
    const issues = this.#spellingIssues;
    pane.total = issues.length;
    pane.entries = issues.map((i) => ({ word: i.word, suggestions: [] }));
    if (this.#spellingActive >= 0 && this.#spellingActive < issues.length) {
      pane.active = this.#spellingActive;
      pane.entries[this.#spellingActive].suggestions = spellSuggestions(
        issues[this.#spellingActive].word,
      );
      pane.entries = [...pane.entries];
    } else {
      pane.active = -1;
    }
  }

  /** Select and scroll to a spelling issue (the pane / command navigation). */
  #gotoSpellingIssue(index: number): void {
    const issues = this.#spellingIssues;
    if (!issues.length) return;
    this.#spellingActive = ((index % issues.length) + issues.length) % issues.length;
    const issue = issues[this.#spellingActive];
    this.editor?.commands.setTextSelection({ from: issue.from, to: issue.to });
    this.#bridge?.scrollIntoView(issue.from);
    this.#syncSpellingPane();
  }

  /** Replace the active issue's text with a suggestion — one transaction, so
   *  undo steps the whole replacement; the re-check rides the render. */
  #replaceSpellingIssue(replacement: string): void {
    const issue = this.#spellingIssues[this.#spellingActive];
    const editor = this.editor;
    if (!issue || !editor) return;
    editor.commands.insertContentAt({ from: issue.from, to: issue.to }, replacement);
    editor.commands.setTextSelection({ from: issue.from, to: issue.from + replacement.length });
  }

  /** Add the active word to the session dictionary, or skip it for this
   *  session — then re-check, which drops the flagged occurrences. */
  #ignoreSpelling(mode: "ignore" | "add"): void {
    const issue = this.#spellingIssues[this.#spellingActive];
    if (!issue) return;
    if (mode === "add") addSpellWord(issue.word);
    else ignoreSpellWord(issue.word);
    this.#runSpellCheck();
  }

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
    const cur = this.#currentSectionProperties()?.pageSize;
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

  /** The current section's sectPr content — the read side of
   *  {@link #updateSectionGeometry}'s write side (same "this section" rule). */
  #currentSectionProperties(): SectionPropertiesOptions | undefined {
    const editor = this.editor;
    if (!editor) return undefined;
    const pos = this.#sectionSectPrPos();
    if (pos != null) {
      const node = editor.state.doc.nodeAt(pos);
      if (!node) return undefined;
      return (node.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
    }
    return (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
  }

  /** Open the Page Setup dialog prefilled from the current section's geometry
   *  in centimeters (the Margins menu's Custom Margins and the Size menu's
   *  More Paper Sizes entries). */
  #openPageSetup(): void {
    const cur = this.#currentSectionProperties();
    // Twips → centimeters for the inputs (2 decimals is Word's display
    // precision); absent geometry — or a UniversalMeasure string form, which
    // the dialog doesn't parse — falls back to Word defaults.
    const cm = (twips?: number | string): number | undefined =>
      typeof twips === "number" ? Math.round(((twips * 2.54) / 1440) * 100) / 100 : undefined;
    // pageMargin/pageSize carry `false` (explicit removal) alongside the
    // properties object — narrow to the object form before reading fields.
    const margin = cur?.pageMargin && typeof cur.pageMargin === "object" ? cur.pageMargin : {};
    const size = cur?.pageSize && typeof cur.pageSize === "object" ? cur.pageSize : {};
    (
      this.shadowRoot?.querySelector("docen-page-setup-dialog") as {
        show(values?: {
          margins?: Partial<PageSetupValues["margins"]>;
          size?: Partial<PageSetupValues["size"]>;
        }): void;
      } | null
    )?.show({
      margins: {
        top: cm(margin.top),
        bottom: cm(margin.bottom),
        left: cm(margin.left),
        right: cm(margin.right),
      },
      size: { width: cm(size.width), height: cm(size.height) },
    });
  }

  /** Open the Columns dialog prefilled from the current section's w:cols
   *  (the Columns menu's More Columns entry). */
  #openColumnsDialog(): void {
    const cur = this.#currentSectionProperties()?.columns;
    // Twips → centimeters for the inputs; absent fields take Word's defaults
    // inside the dialog.
    const columns =
      cur && typeof cur === "object"
        ? cur
        : ({} as Partial<SectionPropertiesOptions["columns"]> & Record<string, unknown>);
    const cm = (twips?: number | string): number | undefined =>
      typeof twips === "number" ? Math.round(((twips * 2.54) / 1440) * 100) / 100 : undefined;
    const raw = columns as {
      count?: number;
      space?: number | string;
      separate?: boolean;
      equalWidth?: boolean;
    };
    (
      this.shadowRoot?.querySelector("docen-columns-dialog") as {
        show(values?: Partial<ColumnsValues>): void;
      } | null
    )?.show({
      count: typeof raw.count === "number" ? raw.count : undefined,
      space: cm(raw.space),
      separate: raw.separate === true,
      equalWidth: raw.equalWidth !== false,
    });
  }

  // The Columns dialog's OK — convert back to twips and write the current
  // section's w:cols. Unequal widths get evenly-split explicit children (the
  // w:col list the projection needs once equalWidth is false); per-column
  // manual widths stay out until the dialog grows inputs for them.
  readonly #onColumnsOk = (event: CustomEvent<ColumnsValues | undefined>): void => {
    const values = event.detail;
    if (!values) return;
    const count = Math.max(1, Math.min(9, Math.trunc(values.count) || 1));
    const space = convertMillimetersToTwip(values.space * 10);
    const children =
      values.equalWidth || count <= 1
        ? undefined
        : Array.from({ length: count }, () => ({
            width: Math.max(
              1,
              Math.floor(((this.#flow?.contentWidthPx ?? 0) * 15 - space * (count - 1)) / count),
            ),
          }));
    this.#mutateCurrentSection((cur) => ({
      ...cur,
      columns: {
        ...cur?.columns,
        count,
        space,
        // Explicit both ways — a conditional spread would let a stale
        // separate:true from the previous w:cols survive an unchecked box.
        separate: values.separate,
        equalWidth: values.equalWidth,
        ...(children ? { children } : {}),
      },
    }));
  };

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

  /** Rewrite the current section's sectPr through `mutate` (Word's "this
   *  section" semantics — a section-carrying paragraph at/after the caret
   *  owns it, otherwise the body-level sectPr) and dispatch. The transaction
   *  re-renders every page of the canvas. */
  #mutateCurrentSection(
    mutate: (cur: SectionPropertiesOptions | undefined) => SectionPropertiesOptions,
  ): void {
    const editor = this.editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const targetPos = this.#sectionSectPrPos();
    if (targetPos != null) {
      const node = doc.nodeAt(targetPos);
      if (node) {
        const cur = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        tr.setNodeMarkup(targetPos, undefined, { ...node.attrs, sectionProperties: mutate(cur) });
      }
    } else {
      const cur = (doc.attrs as { sectionProperties?: SectionPropertiesOptions }).sectionProperties;
      tr.setDocAttribute("sectionProperties", mutate(cur));
    }
    editor.view.dispatch(tr);
  }

  /** Toggle a slot-visibility flag (titlePage / evenAndOddHeaders) on the
   *  current section's sectPr (Word's Different First Page / Odd & Even
   *  Pages). The furniture projection picks the flag up and the page pattern
   *  (first/even slots) follows. */
  #toggleSectionFlag(flag: "titlePage" | "evenAndOddHeaders"): void {
    this.#mutateCurrentSection((cur) => ({
      ...cur,
      [flag]: !(cur as unknown as Record<string, unknown> | undefined)?.[flag],
    }));
  }

  /** Column count for the current section (Word's Page Layout → Columns
   *  presets). The rest of the columns object survives (the gap, the
   *  separator), so toggling back to one column and re-applying keeps the
   *  original geometry. */
  #setColumnCount(count: number): void {
    this.#mutateCurrentSection((cur) => ({
      ...cur,
      columns: { ...cur?.columns, count },
    }));
  }

  /** Line numbering on/off for the current section (w:lnNumType) — Word's
   *  Layout → Line Numbers toggle. */
  #toggleLineNumbers(): void {
    this.#mutateCurrentSection((cur) => ({
      ...cur,
      lineNumberType: cur?.lineNumberType ? undefined : { countBy: 1 },
    }));
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

  /** Resolve a zoom preset to a percent. Numeric presets map directly; the
   *  geometric ones read the stage viewport against the flow box (layout px
   *  at 100%) — page width fills the area width, text width fills it with the
   *  content column, one page fits the whole sheet into the visible height. */
  #zoomPreset(preset: string): void {
    if (/^\d+$/.test(preset)) return this.#setZoom(Number(preset));
    const area = this.shadowRoot?.querySelector("docen-document-area");
    const flow = this.#flow;
    if (!area || !flow) return;
    if (preset === "page-width") return this.#setZoom((area.clientWidth / flow.pageWidthPx) * 100);
    if (preset === "text-width")
      return this.#setZoom((area.clientWidth / flow.contentWidthPx) * 100);
    if (preset === "fit-page") {
      // Whole sheet visible: the net content-box height (clientHeight includes
      // the area's paddings, which would clip the page edges otherwise).
      const style = getComputedStyle(area);
      const visible =
        area.clientHeight -
        Number.parseFloat(style.paddingTop) -
        Number.parseFloat(style.paddingBottom);
      return this.#setZoom(
        Math.min(area.clientWidth / flow.pageWidthPx, visible / flow.pageHeightPx) * 100,
      );
    }
  }

  /** The Zoom dialog (View → Zoom, the status-bar percent click) — prefilled
   *  with the current zoom; the commit applies the preset or free percent. */
  #showZoomDialog(): void {
    (
      this.shadowRoot?.querySelector("docen-zoom-dialog") as { show(zoom: number): void } | null
    )?.show(this.#zoom);
  }

  readonly #onZoomOk = (event: CustomEvent<string | number>): void => {
    if (typeof event.detail === "number") this.#setZoom(event.detail);
    else this.#zoomPreset(event.detail);
  };

  readonly #onZoomOpen = (): void => {
    this.#showZoomDialog();
  };

  /** A status-bar view button (the detail names the status-bar's view:
   *  "reading" | "print" | "web") → the `view` attribute. */
  readonly #onViewSelect = (event: CustomEvent<{ view?: string }>): void => {
    const v = event.detail?.view;
    this.setAttribute("view", v === "reading" ? "read" : v === "web" ? "web" : "print");
  };

  /** Paste Special's pick — re-run the paste in that mode ("text" skips the
   *  rich legs, like the menu's Keep Text Only). */
  readonly #onPasteSpecialOk = (event: CustomEvent<"html" | "text">): void => {
    void this.#paste(event.detail === "text");
  };

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
      bar.setAttribute("view", this.#viewMode());
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

  // The Paragraph dialog's OK — stamp its patch onto every selected paragraph
  // in the editor input currently routes into (a furniture story's editor
  // while a story is open, else the main document).
  readonly #onParagraphOk = (event: CustomEvent<ParagraphDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#bridge?.activeEditor() ?? this.editor;
    target?.commands["paragraph-dialog-apply"]?.(patch);
  };

  // The Font dialog's OK — the patch is the selection's absolute run state
  // (Office commits the dialog atomically): everything lands in ONE chained
  // transaction, so a single undo reverts the whole dialog. (Separate
  // commands can't be fired in sequence off one cached commands object:
  // Tiptap's non-chain commands capture their transaction state once, so a
  // second dispatch applies a stale tr and PM throws "mismatched transaction".)
  readonly #onFontDialogOk = (event: CustomEvent<FontDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#bridge?.activeEditor() ?? this.editor;
    if (!target) return;
    const chain = target.chain();
    // Native attrs ride one textStyle setMark (attrNative null = absent).
    chain.setMark("textStyle", {
      font: patch.font,
      size: patch.size ? Number(patch.size) : null,
      doubleStrike: patch.doubleStrike || null,
      smallCaps: patch.smallCaps || null,
      allCaps: patch.allCaps || null,
      vanish: patch.hidden || null,
    });
    if (patch.bold) chain.setMark("bold");
    else chain.unsetMark("bold");
    if (patch.italic) chain.setMark("italic");
    else chain.unsetMark("italic");
    if (patch.strike) chain.setMark("strike");
    else chain.unsetMark("strike");
    if (patch.underlineStyle) chain["underline-style"](patch.underlineStyle, patch.underlineColor);
    else chain.unsetMark("underline");
    // Sub-/superscript are mutually exclusive marks — commit the checked one
    // and clear both when neither is.
    if (patch.superscript) chain.setMark("superscript");
    else if (patch.subscript) chain.setMark("subscript");
    else {
      chain.unsetMark("superscript");
      chain.unsetMark("subscript");
    }
    chain.run();
    this.#bridge?.focus();
  };

  // The Table Properties dialog's OK — rewrite the caret table's alignment
  // and left indent (the dialog prefills from the same table's attrs).
  readonly #onTablePropertiesOk = (event: CustomEvent<TablePropertiesPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    const target = this.#bridge?.activeEditor() ?? this.editor;
    target?.commands["table-properties-apply"]?.(patch);
    this.#bridge?.focus();
  };

  // The Borders and Shading dialog's OK — route by tab: the border tab
  // stamps the selected paragraphs' w:pBdr, the page tab the current
  // section's w:pgBorders, and the shading tab the paragraph fill.
  readonly #onBordersShadingOk = (event: CustomEvent<BordersDialogPatch | undefined>): void => {
    const patch = event.detail;
    if (!patch) return;
    if (patch.tab === "shading") {
      const target = this.#bridge?.activeEditor() ?? this.editor;
      target?.commands.shading?.(patch.fill ? patch.fill : "none");
      this.#bridge?.focus();
      return;
    }
    if (patch.tab === "border") {
      const target = this.#bridge?.activeEditor() ?? this.editor;
      target?.commands["borders-apply"]?.(patch);
      this.#bridge?.focus();
      return;
    }
    // Page tab — every edge null removes the pgBorders (Word's "none").
    const sides = patch.sides ?? {};
    const edge = (s: BorderSideState | null | undefined): BorderOptions | undefined =>
      s
        ? {
            style: s.style as BorderOptions["style"],
            size: Math.max(2, Math.round(s.size)),
            color: s.color ?? "auto",
            space: 0,
          }
        : undefined;
    const borders: PageBordersOptions | undefined =
      sides.top || sides.bottom || sides.left || sides.right
        ? {
            offsetFrom: "text",
            top: edge(sides.top),
            left: edge(sides.left),
            bottom: edge(sides.bottom),
            right: edge(sides.right),
          }
        : undefined;
    this.#updateSectionGeometry({ pageBorders: borders });
  };

  // Open the Borders and Shading dialog on `tab`, prefilling the border tab
  // from the caret paragraph's w:pBdr and the page tab from the current
  // section's w:pgBorders.
  #openBordersDialog(tab: "border" | "page" | "shading"): void {
    const target = this.#bridge?.activeEditor() ?? this.editor;
    const dialog = this.shadowRoot?.querySelector("docen-borders-shading-dialog") as {
      show(tab: "border" | "page" | "shading", border?: unknown, page?: unknown): void;
    } | null;
    if (!target || !dialog) return;
    // The caret paragraph's attrs (formattable block only — a code block or
    // a table cell still carries paragraph attrs here).
    const { $from } = target.state.selection;
    const block = $from.parent.type.isTextblock
      ? ($from.parent.attrs as Record<string, unknown>)
      : null;
    const border = (block?.border ?? null) as Record<string, unknown> | null;
    const page = (this.#currentSectionProperties()?.pageBorders ?? null) as Record<
      string,
      unknown
    > | null;
    dialog.show(tab, border, page);
  }

  // The Font dialog's prefill — the selection's first text run decides every
  // field (Word reads the same way; a mixed-format selection shows the first
  // run's values). Underline falls back to the textStyle attr channel when no
  // underline mark is present (both carry the same w:u shape).
  #runStateOf(state: EditorState): FontDialogPatch {
    const seen = new Map<string, Record<string, unknown>>();
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
      if (seen.size > 0) return false;
      if (!node.isText) return true;
      for (const m of node.marks)
        if (!seen.has(m.type.name)) seen.set(m.type.name, m.attrs as Record<string, unknown>);
      return false;
    });
    const ts = seen.get("textStyle") ?? {};
    const um = seen.get("underline") as
      | { style?: string | null; color?: string | null }
      | undefined;
    const tsU = ts.underline as { type?: string; color?: string } | undefined;
    const str = (v: unknown): string | null => (typeof v === "string" && v ? v : null);
    return {
      font: str(ts.font),
      size: typeof ts.size === "number" || typeof ts.size === "string" ? String(ts.size) : null,
      bold: seen.has("bold") || ts.bold === true,
      italic: seen.has("italic") || ts.italic === true,
      underlineStyle: um
        ? (str(um.style) ?? "single")
        : tsU && tsU.type && tsU.type !== "none"
          ? tsU.type
          : null,
      underlineColor: um ? str(um.color) : str(tsU?.color),
      strike: seen.has("strike") || ts.strike === true,
      doubleStrike: ts.doubleStrike === true,
      superscript: seen.has("superscript"),
      subscript: seen.has("subscript"),
      smallCaps: ts.smallCaps === true,
      allCaps: ts.allCaps === true,
      hidden: ts.vanish === true,
    };
  }

  // The table grid's pick (hover grid or the classic dialog shape) — the
  // engine's insert-table takes rows/cols (Word's 3×3 preset is the default).
  readonly #onTableInsert = (event: CustomEvent<{ rows?: number; cols?: number }>): void => {
    const { rows, cols } = event.detail ?? {};
    const target = this.#bridge?.activeEditor() ?? this.editor;
    target?.commands["insert-table"]?.({ rows, cols });
  };

  // The Page Setup dialog's OK — convert its centimeters back to twips (the
  // presets go through the same convertMillimetersToTwip) and write the
  // current section's geometry; the transaction re-renders the canvas.
  readonly #onPageSetupOk = (event: CustomEvent<PageSetupValues | undefined>): void => {
    const values = event.detail;
    if (!values) return;
    const twip = (cm: number): number => convertMillimetersToTwip(cm * 10);
    const { margins, size } = values;
    this.#updateSectionGeometry({
      pageMargin: {
        top: twip(margins.top),
        bottom: twip(margins.bottom),
        left: twip(margins.left),
        right: twip(margins.right),
      },
      pageSize: { width: twip(size.width), height: twip(size.height) },
    });
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

  /** Insert → Equation — drop one placeholder template (fraction / script /
   *  radical / sum / integral) at the caret as a math passthrough atom
   *  (Word's Insert → Symbols → Equation gallery). Each argument is an empty
   *  run — the □ slot; the radical's absent degree reads as the square root
   *  (degHide follows). Round-trips verbatim through DOCX; the projection
   *  paints the placeholder box until a math editor lands. */
  #insertEquation(template: string): void {
    const editor = this.editor;
    if (!editor) return;
    const slot = (): object => ({ text: "" });
    const templates: Record<string, object> = {
      fraction: { fraction: { numerator: [slot()], denominator: [slot()] } },
      superScript: { superScript: { children: [slot()], superScript: [slot()] } },
      radical: { radical: { children: [slot()] } },
      sum: {
        sum: {
          children: [slot()],
          subScript: [slot()],
          superScript: [slot()],
          properties: { limitLocation: "undOvr" },
        },
      },
      integral: {
        integral: {
          children: [slot()],
          subScript: [slot()],
          superScript: [slot()],
          properties: { limitLocation: "subSup" },
        },
      },
    };
    const shape = templates[template];
    if (!shape) return;
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify({ math: { children: [shape] } }) },
    };
    const node = editor.schema.nodeFromJSON(seed);
    editor.view.dispatch(editor.state.tr.insert(editor.state.selection.from, node));
  }

  /** Insert → Link / Ctrl+K / right-click Edit Link: open the hyperlink dialog
   *  prefilled from the selection — its text and the link mark riding it (a
   *  caret inside a link edits the whole one via extendMarkRange at commit). */
  #insertLink(): void {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    const { empty, from, to } = editor.state.selection;
    (
      this.shadowRoot?.querySelector("docen-link-dialog") as {
        show(values?: Partial<LinkValues>): void;
      } | null
    )?.show({
      text: empty ? "" : editor.state.doc.textBetween(from, to, " "),
      href: editor.getAttributes("link").href as string | undefined,
    });
  }

  // The Link dialog's OK — Word's Insert Link semantics: an empty address
  // removes an existing link; a selection gets marked (its text replaced when
  // the dialog's display text was edited); an empty selection inserts fresh
  // display text carrying the mark.
  readonly #onLinkOk = (event: CustomEvent<LinkValues | undefined>): void => {
    const values = event.detail;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!values || !editor) return;
    this.#bridge?.focus();
    const raw = values.href.trim();
    // The link mark riding the selection, if any.
    const existing = editor.getAttributes("link").href as string | undefined;
    if (raw === "") {
      if (existing) editor.chain().focus().extendMarkRange("link").unsetLink().run();
      return;
    }
    // `#name` stays a bookmark anchor; bare hosts gain the https scheme.
    const href = raw.startsWith("#") || /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
    const mark = [
      { type: "link", attrs: { href, target: href.startsWith("#") ? null : "_blank" } },
      { type: "textStyle", attrs: { style: "Hyperlink" } },
    ] as const;
    const { empty, from, to } = editor.state.selection;
    if (!empty) {
      const text = values.text.trim();
      const selected = editor.state.doc.textBetween(from, to, " ");
      if (text && text !== selected) {
        // The display text was edited — replace the selection with the fresh
        // marked run (one undo step).
        editor.commands.insertContentAt({ from, to }, { type: "text", text, marks: [...mark] });
        return;
      }
      // Word stamps hyperlink runs with the "Hyperlink" character style —
      // that style (not the w:hyperlink element) paints links blue.
      editor.chain().focus().extendMarkRange("link").setLink({ href }).run();
      editor.commands.setMark("textStyle", { style: "Hyperlink" });
      return;
    }
    // No selection: the display text inserts marked.
    const text = values.text.trim();
    if (!text) return;
    editor.commands.insertContent({ type: "text", text, marks: [...mark] });
  };

  /** The href of the link mark at the caret (the context menu's open/copy
   *  source; the right-click collapsed the caret onto the link first), or
   *  null. */
  #hrefAtCaret(): string | null {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return null;
    const mark = editor.state.doc
      .resolve(Math.min(editor.state.selection.from, editor.state.doc.content.size))
      .marks()
      .find((m) => m.type.name === "link");
    const href = mark?.attrs.href;
    return typeof href === "string" && href ? href : null;
  }

  /** Ctrl+Click / Open Hyperlink on a `#name` link — place the caret past the
   *  matching bookmarkStart atom and scroll it into view (Word scrolls to the
   *  bookmark). No matching bookmark is a no-op. */
  #jumpToBookmark(name: string): void {
    const editor = this.editor;
    if (!editor) return;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (target != null || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
          bookmarkStart?: { name?: string };
        };
        if (data.bookmarkStart?.name === name) target = pos + child.nodeSize;
      } catch {
        // opaque verbatim blob — not a bookmark
      }
    });
    if (target == null) return;
    this.#setTextSelection(target);
    this.#bridge?.scrollIntoView(target);
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
    if (target != null) {
      this.#setTextSelection(target + 1);
      this.#bridge?.scrollIntoView(target + 1);
    }
  }

  /** References → Previous Footnote: place the caret on the previous
   *  footnote/endnote reference before the selection (document order). */
  #jumpPreviousNote(): void {
    const editor = this.editor;
    if (!editor) return;
    const { from } = editor.state.selection;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (pos >= from || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as Record<string, unknown>;
        if ("footnoteReference" in data || "endnoteReference" in data) target = pos;
      } catch {
        // opaque verbatim blob — not a note reference
      }
    });
    if (target != null) {
      this.#setTextSelection(target + 1);
      this.#bridge?.scrollIntoView(target + 1);
    }
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

  /** Right-click on the canvas — Word's context menu, rebuilt per click.
   *  Clicking outside the selection first moves the caret there (Word's
   *  behavior), the clipboard section appears only with a selection, and a
   *  click on a hyperlink swaps the Link item for Edit/Remove. Menu items
   *  dispatch the same command ids as the ribbon, so #onCommand handles them.
   *  While a furniture story (header/footer) is being edited the positions
   *  belong to the story's editor, which these main-story commands cannot
   *  target — suppress the menu there. */
  readonly #onContextMenu = (event: MouseEvent): void => {
    const menu = this.shadowRoot?.querySelector("docen-context-menu") ?? null;
    const editor = this.editor;
    if (!menu || !editor || !event.composedPath().includes(menu)) return;
    if (this.#bridge?.storyKind() != null) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const { selection } = editor.state;
    const pos = this.#bridge?.posAtClient(event.clientX, event.clientY) ?? null;
    const inSelection =
      !selection.empty && pos != null && pos >= selection.from && pos <= selection.to;
    const onLink =
      pos != null &&
      editor.state.doc
        .resolve(Math.min(Math.max(pos, 0), editor.state.doc.content.size))
        .marks()
        .some((m) => m.type.name === "link");
    // The click sits inside a table when any ancestor (nearest wins for the
    // command) is a table node — Word then carries table entries on the menu.
    const inTable =
      pos != null &&
      (() => {
        const $p = editor.state.doc.resolve(
          Math.min(Math.max(pos, 0), editor.state.doc.content.size),
        );
        for (let d = $p.depth; d > 0; d -= 1) {
          if ($p.node(d).type === editor.state.schema.nodes.table) return true;
        }
        return false;
      })();
    // Word: a right-click outside the selection collapses the caret there.
    if (pos != null && !inSelection) editor.commands.setTextSelection(pos);
    const items: RibbonMenuItem[] = [];
    if (inSelection) {
      items.push({ text: t("context.cut", this), event: "cut" });
      items.push({ text: t("context.copy", this), event: "copy" });
    }
    items.push({ text: t("context.paste", this), event: "paste" });
    items.push({
      text: t("context.keep-text-only", this),
      event: "paste",
      value: "keep-text-only",
    });
    items.push({ text: "-" });
    if (onLink) {
      items.push({ text: t("context.open-link", this), event: "open-link" });
      items.push({ text: t("context.copy-link", this), event: "copy-link" });
      items.push({ text: t("context.edit-link", this), event: "link" });
      items.push({ text: t("context.unlink", this), event: "unset-link" });
      items.push({ text: "-" });
      items.push({ text: t("context.comment", this), event: "new-comment" });
      items.push({ text: "-" });
    } else if (inSelection) {
      items.push({ text: t("context.link", this), event: "link" });
      items.push({ text: t("context.comment", this), event: "new-comment" });
      items.push({ text: "-" });
    }
    items.push({ text: t("context.select-all", this), event: "select" });
    if (inTable) {
      items.push({ text: "-" });
      items.push({ text: t("ribbon.cmd.insert-row-above", this), event: "insert-row-above" });
      items.push({ text: t("ribbon.cmd.insert-row-below", this), event: "insert-row-below" });
      items.push({ text: t("ribbon.cmd.insert-column-left", this), event: "insert-column-left" });
      items.push({ text: t("ribbon.cmd.insert-column-right", this), event: "insert-column-right" });
      items.push({ text: "-" });
      items.push({ text: t("ribbon.cmd.delete-row", this), event: "delete-row" });
      items.push({ text: t("ribbon.cmd.delete-column", this), event: "delete-column" });
      items.push({ text: t("context.delete-table", this), event: "delete-table" });
      // Word's merge/split + AutoFit + the Properties entry close the table
      // menu (commands shared with the Table Layout tab).
      items.push({ text: "-" });
      items.push({ text: t("ribbon.cmd.merge-cells", this), event: "merge-cells" });
      items.push({ text: t("ribbon.cmd.split-cell", this), event: "split-cell" });
      items.push({ text: t("ribbon.opt.autofit-contents", this), event: "autofit-contents" });
      items.push({ text: t("ribbon.opt.autofit-window", this), event: "autofit-window" });
      items.push({ text: "-" });
      items.push({ text: t("context.table-properties", this), event: "table-properties" });
    }
    menu.setAttribute("items", JSON.stringify(items));
  };

  /** Review → New Comment: open the floating compose box beside the
   *  selection (Word's Simple-Markup reply card — it hangs in the margin at
   *  the anchored line and scrolls with its page). The text arrives via the
   *  `comment:create` event (#onCommentCreate commits it). Without a
   *  selection the word at the caret anchors the comment (Word for the web's
   *  behavior); a caret on whitespace is a no-op. */
  #insertComment(): void {
    const editor = this.editor;
    if (!editor) return;
    // Spread would miss from/to — they're prototype getters on Selection.
    const { from, to } = editor.state.selection;
    if (from === to) {
      const word = wordRangeAt(editor.state.doc, from);
      if (!word) return;
      this.#pendingCommentRange = word;
    } else {
      this.#pendingCommentRange = { from, to };
    }
    this.#openCommentCompose();
  }

  /** The floating compose box over the canvas (null once dismissed). */
  #commentCompose?: HTMLElement;

  /** Mount the floating compose card on the anchored page frame — Fluent
   *  components (text-area, buttons) over the design-token palette, no
   *  hand-rolled colors. */
  #openCommentCompose(): void {
    this.#closeCommentCompose();
    const range = this.#pendingCommentRange;
    if (!range) return;
    const anchor = this.#bridge?.commentAnchorRect(range.from, range.to);
    if (!anchor) {
      // Unmappable selection (stale map) — drop the pending compose.
      this.#pendingCommentRange = undefined;
      return;
    }
    const card = document.createElement("div");
    card.setAttribute("data-docen-overlay", "");
    Object.assign(card.style, {
      position: "absolute",
      zIndex: "6",
      width: "220px",
      boxSizing: "border-box",
      padding: "10px",
      display: "flex",
      flexDirection: "column",
      gap: "8px",
      background: "var(--docen-color-bg, #ffffff)",
      border: "1px solid var(--docen-color-divider, #e2e2e2)",
      borderRadius: "var(--borderRadiusLarge, 8px)",
      boxShadow: "var(--shadow4, 0 4px 8px rgba(0,0,0,.14))",
      fontFamily: "inherit",
      fontSize: "var(--docen-font-size-ribbon, 12px)",
    } satisfies Partial<CSSStyleDeclaration>);
    const area = document.createElement("fluent-textarea") as HTMLTextAreaElement & HTMLElement;
    // `block` drops Fluent's fixed 18rem inline-size — without it the inner
    // root box overflows the 220px card.
    area.setAttribute("block", "");
    area.setAttribute("resize", "vertical");
    area.setAttribute("rows", "3");
    area.setAttribute("placeholder", t("comments.placeholder", this));
    const row = document.createElement("div");
    Object.assign(row.style, {
      display: "flex",
      gap: "6px",
      justifyContent: "flex-end",
    } satisfies Partial<CSSStyleDeclaration>);
    const cancel = document.createElement("fluent-button");
    cancel.setAttribute("appearance", "neutral");
    cancel.textContent = t("comments.cancel", this);
    const post = document.createElement("fluent-button");
    post.setAttribute("appearance", "accent");
    post.textContent = t("comments.post", this);
    cancel.addEventListener("click", () =>
      this.dispatchEvent(new CustomEvent("comment:cancel", { bubbles: true, composed: true })),
    );
    const postIt = (): void => {
      const text = (area.value ?? "").trim();
      if (!text) return;
      this.dispatchEvent(
        new CustomEvent("comment:create", { bubbles: true, composed: true, detail: { text } }),
      );
    };
    post.addEventListener("click", postIt);
    // Ctrl+Enter posts (the sidebar edit's shortcut); Escape cancels.
    area.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        postIt();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        this.dispatchEvent(new CustomEvent("comment:cancel", { bubbles: true, composed: true }));
      }
    });
    row.append(cancel, post);
    card.append(area, row);
    // Word hangs the card in the margin just past the anchored line's end,
    // clamped into the page frame when the line runs to the right edge.
    const frameW = anchor.frame.clientWidth;
    const CARD_W = 220;
    const left = Math.min(anchor.left + 12, Math.max(frameW - CARD_W - 8, 0));
    Object.assign(card.style, {
      left: `${left}px`,
      top: `${anchor.top}px`,
    } satisfies Partial<CSSStyleDeclaration>);
    anchor.frame.append(card);
    this.#commentCompose = card;
    // The host is not focusable — focus lands on the shadow <textarea>, or
    // every keystroke goes to the canvas bridge and into the document.
    requestAnimationFrame(() => {
      const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
      input?.focus();
    });
  }

  /** Tear down the floating compose card (Post, Cancel, or Escape). */
  #closeCommentCompose(): void {
    this.#commentCompose?.remove();
    this.#commentCompose = undefined;
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
    this.#closeCommentCompose();
    if (!editor || !text || !range) return;
    // Word returns the caret to the body once the comment is committed.
    this.#bridge?.focus();
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
    this.#closeCommentCompose();
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
  #commentsPaneEl(): (HTMLElement & { comments?: string; activeId?: string }) | null {
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
   *  every transaction (the pane is a pure view of the model). Cards order by
   *  their anchored range's document position — Word's sidebar follows the
   *  text, not the round-trip append order the extras array stores. */
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
    const startPos = new Map<number, number>();
    this.editor?.state.doc.descendants((child, pos) => {
      const marker = DocenDocument.commentMarkerOf(child);
      if (marker?.kind === "start") startPos.set(marker.id, pos);
    });
    const cards = (docAttrs.documentExtras?.comments ?? [])
      .map((c) => ({
        id: Number(c.id ?? 0),
        author: c.author ?? "",
        initials: c.initials ?? "",
        date: c.date ?? "",
        text: (c.children ?? []).map((r) => r.text ?? "").join(""),
        pos: startPos.get(Number(c.id ?? 0)),
      }))
      .sort((a, b) => (a.pos ?? Number.MAX_SAFE_INTEGER) - (b.pos ?? Number.MAX_SAFE_INTEGER))
      .map(({ pos: _pos, ...card }) => card);
    pane.comments = JSON.stringify(cards);
  }

  /** selectionUpdate → highlight the comments-pane card whose anchored range
   *  covers the selection (Word paints the anchored card while the caret sits
   *  in its text). */
  readonly #syncActiveCommentCard = (): void => {
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    const id = this.#activeCommentId();
    pane.setAttribute("active-id", id == null ? "" : String(id));
  };

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

  /** Design → Paragraph Spacing presets — stamp the styles' docDefaults
   *  paragraph spacing (styles.default.document.paragraph.spacing), the
   *  document-level default every paragraph without explicit spacing
   *  inherits. Word's preset values: default restores the factory 8pt-after /
   *  1.08-line spacing; the named presets are single-spaced with the after
   *  gap shrinking (none 0pt → compact 2pt → narrow 6pt → wide 16pt). */
  #setParagraphSpacing(preset?: string): void {
    const editor = this.editor;
    if (!editor) return;
    const spacing =
      preset === "none"
        ? { before: 0, after: 0, line: 240, lineRule: "auto" }
        : preset === "compact"
          ? { after: 40, line: 240, lineRule: "auto" }
          : preset === "narrow"
            ? { after: 120, line: 240, lineRule: "auto" }
            : preset === "wide"
              ? { after: 320, line: 240, lineRule: "auto" }
              : preset === "default"
                ? { after: 160, line: 259, lineRule: "auto" }
                : null;
    if (!spacing) return;
    const styles = { ...((editor.state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const defaults = { ...((styles.default ?? {}) as Record<string, unknown>) };
    const documentDefaults = { ...((defaults.document ?? {}) as Record<string, unknown>) };
    documentDefaults.paragraph = {
      ...((documentDefaults.paragraph ?? {}) as Record<string, unknown>),
      spacing,
    };
    defaults.document = documentDefaults;
    styles.default = defaults;
    editor.view.dispatch(editor.state.tr.setDocAttribute("styles", styles));
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

  /** Design → Watermark presets (Word's gallery): one diagonal silver text
   *  shape stamped into every header slot — Word's watermark IS a
   *  behind-document, page-centered shape anchored in the header, so it
   *  repeats on every page of the slot. The shape carries Word's watermark
   *  name ("WordPictureWatermark"), which is also how Remove finds it. */
  #setWatermark(preset?: string): void {
    const spec = preset && preset !== "remove" ? WATERMARK_PRESETS[preset] : undefined;
    this.#stampWatermark(spec ? watermarkPara(spec) : null);
  }

  /** Strip any existing watermark from every section's header slots, then
   *  append the given stamp paragraph (null removes). Word's watermark rides
   *  linked headers and reads on every page, so every section's carrier is
   *  stamped — earlier sections' slots live on their closing sectPr
   *  paragraphs, the final section's on the doc node. Shared by the gallery
   *  presets and the custom dialog's text/picture stamps. */
  #stampWatermark(para: JSONContent | null): void {
    const editor = this.editor;
    if (!editor) return;
    const { doc, tr } = editor.state;
    const stamp = (attrs: Record<string, unknown>): Record<string, unknown> => {
      const stamped = stampHeaderSlots(
        attrs.sectionHeaders as Record<string, JSONContent[] | undefined>,
        para,
      );
      // An all-empty sectionHeaders is the no-headers state (Word's Remove
      // Watermark leaves a blank header behind; an empty attrs object is the
      // cleaner equivalent here and drops the furniture strut).
      const anyContent = Object.values(stamped).some((paras) => paras.length > 0);
      return { ...attrs, sectionHeaders: anyContent ? stamped : {} };
    };
    doc.descendants((node, pos) => {
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        tr.setNodeMarkup(pos, undefined, stamp(node.attrs as Record<string, unknown>));
      }
      return true;
    });
    tr.setDocAttribute(
      "sectionHeaders",
      stamp(doc.attrs as Record<string, unknown>).sectionHeaders,
    );
    editor.view.dispatch(tr);
  }

  /** The custom watermark dialog's OK (Word's 自定义水印): none clears, the
   *  text spec stamps the text shape, the picture spec probes the natural
   *  size then stamps the floating picture. */
  readonly #onWatermarkOk = (
    event: CustomEvent<
      | { kind: "none" }
      | { kind: "text"; spec: WatermarkTextSpec }
      | { kind: "picture"; spec: WatermarkPictureSpec }
      | undefined
    >,
  ): void => {
    const detail = event.detail;
    if (!detail) return;
    if (detail.kind === "none") {
      this.#stampWatermark(null);
    } else if (detail.kind === "text") {
      this.#stampWatermark(customTextWatermarkPara(detail.spec));
    } else {
      void probeImageSize(detail.spec.src).then((natural) => {
        this.#stampWatermark(pictureWatermarkPara(detail.spec, natural));
      });
    }
    this.#bridge?.focus();
  };

  /** Design → Page Borders presets — stamp w:pgBorders on the current
   *  section (Word's Borders and Shading gallery): none clears it; box is a
   *  plain rule; shadow thickens the bottom/right edges; double and dashed
   *  swap the rule's style. Sides measure from the text margin (Word's
   *  default offsetFrom), 0.5 pt black. */
  #setPageBorders(preset?: string): void {
    if (!preset) return;
    const side = (style: BorderOptions["style"], size = 4): BorderOptions => ({
      style,
      size,
      space: 0,
    });
    const rule: BorderOptions["style"] =
      preset === "double" ? "double" : preset === "dashed" ? "dashSmallGap" : "single";
    const borders: PageBordersOptions | undefined =
      preset === "none"
        ? undefined
        : preset === "shadow"
          ? {
              offsetFrom: "text",
              top: side("single"),
              left: side("single"),
              bottom: side("single", 18),
              right: side("single", 18),
            }
          : {
              offsetFrom: "text",
              top: side(rule),
              right: side(rule),
              bottom: side(rule),
              left: side(rule),
            };
    // pageBorders rides the top-level spread in mergeSectionProperties (an
    // undefined patch value removes the pgBorders — Word's "none").
    this.#updateSectionGeometry({ pageBorders: borders });
  }

  /** Design → Watermark → Custom Watermark: open the dialog prefilled from
   *  the current stamp (a text shape's run reads back text/color/size;
   *  a picture stamp selects the picture pane). */
  #openWatermarkDialog(): void {
    const dialog = this.shadowRoot?.querySelector("docen-watermark-dialog") as {
      show(current?: unknown): void;
    } | null;
    const editor = this.editor;
    if (!dialog) return;
    if (!editor) {
      dialog.show();
      return;
    }
    // The stamp rides every section's carrier (see #stampWatermark) — read
    // them in document order and prefill from the first stamp found.
    const groups: Array<Record<string, JSONContent[] | undefined>> = [];
    editor.state.doc.descendants((node) => {
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        groups.push(
          (node.attrs as { sectionHeaders?: Record<string, JSONContent[] | undefined> })
            .sectionHeaders ?? {},
        );
      }
      return true;
    });
    groups.push(
      (editor.state.doc.attrs.sectionHeaders ?? {}) as Record<string, JSONContent[] | undefined>,
    );
    let current: unknown = null;
    for (const headers of groups) {
      current = this.#watermarkSpecOf(headers);
      if (current) break;
    }
    dialog.show(current);
  }

  /** One section's header slots read back as the dialog's prefill — the first
   *  watermark shape/picture in the default slot, null when none. */
  #watermarkSpecOf(headers: Record<string, JSONContent[] | undefined>): unknown {
    for (const para of headers.default ?? []) {
      for (const child of ((para as JSONContent).content ?? []) as JSONContent[]) {
        if (child.type === "wpsShape" && isWatermarkNode(child)) {
          const shape = (child.attrs as { wpsShape?: Record<string, unknown> }).wpsShape ?? {};
          const run = ((child as JSONContent).content?.[0]?.content ?? []) as Array<{
            marks?: Array<{ type: string; attrs: Record<string, unknown> }>;
          }>;
          const style = run[0]?.marks?.find((m) => m.type === "textStyle")?.attrs ?? {};
          return {
            kind: "text",
            text: (child.content?.[0]?.content?.[0] as { text?: string })?.text ?? "",
            font: (style.font as string) ?? null,
            size: (style.size as number) ?? null,
            color: (style.color as string) ?? "C0C0C0",
            diagonal:
              (shape.transformation as { rotation?: number })?.rotation != null &&
              (shape.transformation as { rotation?: number }).rotation! < 0,
            semiTransparent: false,
          };
        }
        if (child.type === "image" && isWatermarkNode(child)) {
          const attrs = child.attrs as { blipEffects?: { luminance?: unknown } };
          return {
            kind: "picture",
            hasImage: true,
            washout: !!attrs.blipEffects?.luminance,
            scale: "auto",
          };
        }
      }
    }
    return null;
  }

  /** Insert → Text Box / Shapes: a standalone wps shape run, floating
   *  wrap-none and centered on the page (Word's insertion behavior). The
   *  text box carries Word's plain look — white fill, accent-1 hairline —
   *  and an editable empty body (the PM `content`); a gallery shape carries
   *  its preset geometry with the accent fill instead. */
  #insertShape(preset: string | undefined): void {
    // Insert into the story the caret lives in — a header/footer story must
    // receive the shape, not the stale main-doc selection behind it.
    const editor = this.#bridge?.activeEditor() ?? this.editor;
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
      geometry.geometry = preset;
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
    // Read-only documents (Viewing mode) reject document-changing commands —
    // the viewless editor has no DOM surface to refuse them, so the gate
    // lives here (Word's read-only ribbon). Chrome actions and clipboard
    // reads stay live.
    if (this.editor && !this.editor.isEditable && !READONLY_LIVE.has(name)) {
      return;
    }
    // UI chrome actions are handled locally and need no Tiptap editor.
    if (name === "toggle-navigation") {
      this.#togglePane("navigation");
      return;
    }
    // View → Outline: Word's outline view maps to the document-structure
    // pane here (the same tree the navigation pane shows).
    if (name === "outline") {
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
    // The Table button's face opens the hover grid; its dropdown's Insert
    // Table opens the classic dialog shape (both insert via table-grid:insert).
    if (name === "insert-table") {
      (
        this.shadowRoot?.querySelector("docen-table-dialog") as { show(m?: string): void } | null
      )?.show("grid");
      return;
    }
    if (name === "table-dialog") {
      (
        this.shadowRoot?.querySelector("docen-table-dialog") as { show(m?: string): void } | null
      )?.show("form");
      return;
    }
    // Page setup actions write sectionProperties; the transaction re-renders.
    // "more"/"custom" open the Page Setup dialog instead of a preset.
    if (name === "page-size") {
      if (value === "more") this.#openPageSetup();
      else this.#setPageSize(value);
      return;
    }
    if (name === "orientation") {
      this.#setOrientation(value);
      return;
    }
    if (name === "margins") {
      if (value === "custom") this.#openPageSetup();
      else this.#setMargins(value);
      return;
    }
    // Columns presets (the Layout tab's Columns menu: one/two/three);
    // More Columns opens the dialog prefilled from the current section.
    if (name === "columns") {
      if (value === "more") this.#openColumnsDialog();
      else {
        const count = Number(value);
        if (count >= 1 && count <= 9) this.#setColumnCount(count);
      }
      return;
    }
    // Line Numbers toggle (the Layout tab's Line Numbers button).
    if (name === "line-numbers") {
      this.#toggleLineNumbers();
      return;
    }
    // AutoFit Window needs the page's text width — a layout value the command
    // layer can't see, so the host injects it as the twip value (px × 15 at
    // the layout's 96 dpi).
    if (name === "autofit-window") {
      const flow = this.#flow;
      const ed = this.editor;
      if (ed && flow && flow.contentWidthPx > 0) {
        (ed.commands as unknown as Record<string, (v?: string) => unknown>)["autofit-window"](
          String(Math.round(flow.contentWidthPx * 15)),
        );
      }
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
      if (value === "zoom-dialog") this.#showZoomDialog();
      else if (value) this.#zoomPreset(value);
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
      // Insert/update in the story the caret lives in (a header/footer story
      // opening must not send the TOC into the stale main-doc selection).
      const target = this.#bridge?.activeEditor() ?? editor;
      const pageOf = (pos: number): number | null => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const tabPositionTw = this.#flow
        ? Math.round(this.#flow.contentWidthPx / twipToPx(1))
        : undefined;
      const ran = target.commands[name](pageOf, tabPositionTw);
      if (name === "toc" && ran) {
        // Frame N re-flows (the bridge's raf-merged onDoc), frame N+1 the
        // caret map carries the post-insert pagination.
        requestAnimationFrame(() =>
          requestAnimationFrame(() => target.commands["update-toc"](pageOf, tabPositionTw)),
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
    // The Header & Footer context tab — switch stories (the dirty close rides
    // the normal exit path), flip the same slot flags, and close.
    if (name === "goto-header" || name === "goto-footer") {
      const page =
        this.#storyPage >= 0 ? this.#storyPage : this.#bridge?.pageOf(editor.state.selection.from);
      this.#bridge?.exitStory();
      if (page != null)
        this.#bridge?.enterStory(name === "goto-header" ? "header" : "footer", page);
      return;
    }
    if (name === "close-header-footer") {
      this.#bridge?.exitStory();
      return;
    }
    if (name === "header-option") {
      if (value === "title-page" || value === "odd-even")
        this.#toggleSectionFlag(value === "title-page" ? "titlePage" : "evenAndOddHeaders");
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
    // Paragraph — open the dialog prefilled from the caret paragraph's attrs;
    // the commit arrives via paragraph:ok (stamped by paragraph-dialog-apply).
    if (name === "paragraph-dialog") {
      const target = this.#bridge?.activeEditor() ?? editor;
      const node = target?.state.selection.$from.parent;
      if (node?.type.name === "paragraph") {
        (
          this.shadowRoot?.querySelector("docen-paragraph-dialog") as {
            show(attrs?: Record<string, unknown>): void;
          } | null
        )?.show(node.attrs as Record<string, unknown>);
      }
      return;
    }
    // Font — open the dialog prefilled from the selection's run marks; the
    // commit arrives via font:ok (#onFontDialogOk).
    if (name === "font-dialog") {
      const target = this.#bridge?.activeEditor() ?? editor;
      const dialog = this.shadowRoot?.querySelector("docen-font-dialog") as {
        show(state: FontDialogPatch): void;
      } | null;
      if (target && dialog) dialog.show(this.#runStateOf(target.state));
      return;
    }
    // Table Properties — open the dialog prefilled from the caret table's
    // attrs; the commit arrives via table-properties:ok
    // (table-properties-apply). No caret table → nothing to show.
    if (name === "table-properties") {
      const target = this.#bridge?.activeEditor() ?? editor;
      const anchor = target ? tableAncestry(target.state) : null;
      const dialog = this.shadowRoot?.querySelector("docen-table-properties-dialog") as {
        show(attrs?: Record<string, unknown>): void;
      } | null;
      if (target && anchor && dialog) {
        dialog.show(
          target.state.selection.$from.node(anchor.tableAt).attrs as Record<string, unknown>,
        );
      }
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
      else if (value === "prev") this.#jumpPreviousNote();
      else this.#insertNote("footnote");
      return;
    }
    // Equation — insert one placeholder math template at the caret (Word's
    // Insert → Symbols → Equation gallery).
    if (name === "equation") {
      this.#insertEquation(String(value));
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
    // The Borders and Shading dialog entries — the border split's and the
    // page-border split's last item carry the dialog value; the source split
    // picks the tab (the remaining preset values fall through below).
    if (value === "borders-shading" && (name === "border" || name === "page-border")) {
      this.#openBordersDialog(name === "page-border" ? "page" : "border");
      return;
    }
    if (name === "page-border") {
      this.#setPageBorders(value);
      return;
    }
    // Paragraph Spacing presets — stamp the styles' docDefaults paragraph
    // spacing (Word's Design → Paragraph Spacing; the document-level default
    // every paragraph without explicit spacing inherits).
    if (name === "paragraph-spacing") {
      this.#setParagraphSpacing(typeof value === "string" ? value : undefined);
      return;
    }
    // View toggles — ruler and gridlines are paint-time view state (never in
    // the document), so the stage flips the flag and repaints.
    if (name === "toggle-ruler") {
      this.#stage?.setShowRuler(!this.#stage.showRuler);
      return;
    }
    if (name === "toggle-gridlines") {
      this.#stage?.setShowGridlines(!this.#stage.showGridlines);
      return;
    }
    // Watermark gallery — a preset id stamps the header shape, "remove"
    // strips it; the custom entry opens Word's watermark dialog (Word's
    // Design → Watermark split button).
    if (name === "watermark") {
      if (value === "custom") {
        this.#openWatermarkDialog();
        return;
      }
      this.#setWatermark(typeof value === "string" ? value : undefined);
      return;
    }
    // Link — prompt for an address and mark the selection (or insert fresh
    // display text when the selection is empty).
    if (name === "link") {
      this.#insertLink();
      return;
    }
    // Context menu → Remove Hyperlink: unset the link mark across the
    // right-clicked link (extendMarkRange reaches past the caret's spot).
    if (name === "unset-link") {
      // The link mark spans the right-clicked range in whichever editor the
      // caret lives in (a furniture story has its own links).
      (this.#bridge?.activeEditor() ?? this.editor)
        ?.chain()
        .extendMarkRange("link")
        .unsetLink()
        .run();
      return;
    }
    // Context menu → Open Hyperlink: `#name` jumps to its bookmark, anything
    // else opens in a new window.
    if (name === "open-link") {
      const href = this.#hrefAtCaret();
      if (!href) return;
      if (href.startsWith("#")) this.#jumpToBookmark(href.slice(1));
      else window.open(href, "_blank", "noopener,noreferrer");
      return;
    }
    // Context menu → Copy Hyperlink: the address to the system clipboard.
    if (name === "copy-link") {
      const href = this.#hrefAtCaret();
      if (href) void navigator.clipboard.writeText(href);
      return;
    }
    // New Comment — anchor the selection (or the word at the caret) with a
    // Word comment; Edit/Delete operate on the comment covering the selection.
    if (name === "new-comment") {
      this.#insertComment();
      return;
    }
    // The trailing title-bar "comment" button toggles the comments pane
    // (Word's sidebar) — it lists every comment, it does not create one.
    if (name === "comment") {
      this.#togglePane("comments");
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
    // so copy/cut route through the bridge's lane (it pins the slice payload
    // exactly like a keyboard copy, keeping every paste entry lossless).
    if (name === "copy" || name === "cut") {
      void this.#bridge?.copySelection(name === "cut");
      return;
    }
    if (name === "paste") {
      if (value === "paste-special") {
        (
          this.shadowRoot?.querySelector("docen-paste-special-dialog") as unknown as {
            show(): void;
          } | null
        )?.show();
        return;
      }
      void this.#paste(value === "keep-text-only");
      return;
    }
    // Home → Clipboard group launcher — the Office Clipboard pane.
    if (name === "clipboard-dialog") {
      this.#togglePane("clipboard");
      return;
    }
    // View → the four view buttons (Word's View tab): each selects a document
    // view through the `view` attribute — #applyView restages the render.
    const viewOf: Record<string, string> = {
      "print-layout": "print",
      "web-layout": "web",
      "read-mode": "read",
      draft: "draft",
    };
    if (viewOf[name]) {
      this.setAttribute("view", viewOf[name]);
      return;
    }
    // Spelling (Review → Spelling & Grammar, F7, the status-bar book):
    // re-check now, open the pane, and start at the first issue at/after the
    // caret (Word starts checking from the insertion point).
    if (name === "spell-check") {
      this.#runSpellCheck();
      this.#setTaskpane("proofing", true);
      const from = this.editor?.state.selection.from ?? 0;
      const issues = this.#spellingIssues;
      if (issues.length) {
        const first = issues.find((issue) => issue.from >= from) ?? issues[0];
        this.#gotoSpellingIssue(issues.indexOf(first));
      }
      return;
    }
    // Language (Review → Language, the status-bar language item): the
    // proofing-language dialog for the selection.
    if (name === "language") {
      this.#onLanguageOpen();
      return;
    }
    // Phonetic guide (拼音指南, Home → Font): the per-character reading
    // dialog over the selection.
    if (name === "phonetic-guide") {
      this.#onPhoneticOpen();
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
    // They target the editor input currently routes into — while a furniture
    // story is open that's the story's editor, not the main document (whose
    // selection is stale and would be stamped instead).
    const target = this.#bridge?.activeEditor() ?? editor;
    const commands = target.commands as unknown as Record<string, (value?: string) => unknown>;
    const cmd = commands[name];
    if (typeof cmd === "function") {
      cmd(value);
      if (name === "next-change" || name === "previous-change") {
        this.#bridge?.scrollIntoView(target.state.selection.from);
      }
      // The comboboxes keep focus to filter their lists — after a pick, hand
      // the keyboard back to the document (buttons never take it: their
      // mousedown preventDefaults).
      if (name === "font-name" || name === "font-size" || name === "style") {
        this.#bridge?.focus();
      }
      return;
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

  /** The caret's proofing language (the textStyle mark's w:lang fields).
   *  Runs without an explicit mark show the default proofing language —
   *  Word mirrors the editing language implied by the UI locale here. */
  #caretLanguage(): { value: string; noProof: boolean } {
    const editor = this.editor;
    const mark = editor?.state.selection.$from.marks().find((m) => m.type.name === "textStyle");
    const language = mark?.attrs.language as { value?: string } | undefined;
    const fallback = (document.documentElement.lang || "en").startsWith("zh") ? "zh-CN" : "en-US";
    return { value: language?.value || fallback, noProof: mark?.attrs.noProof === true };
  }

  /** Status-bar language item / Review → Language — open the dialog prefilled
   *  from the caret's current proofing language. */
  readonly #onLanguageOpen = (): void => {
    const dialog = this.shadowRoot?.querySelector("docen-language-dialog") as unknown as {
      show(tag: string | null, noProof?: boolean): void;
    } | null;
    const { value, noProof } = this.#caretLanguage();
    dialog?.show(value || null, noProof);
  };

  /** Language dialog 确定 — stamp the proofing language onto the selection's
   *  runs (w:lang), plus/minus the "do not check spelling" flag (w:noProof).
   *  Like Word, an empty selection is a no-op here (no input-language service
   *  to feed). One chained transaction (the font-dialog pattern): `focus()`
   *  resets the viewless editor's selection, so the bridge restores it after. */
  readonly #onLanguageOk = (event: Event): void => {
    const { value, noProof } = (event as CustomEvent<{ value?: string; noProof?: boolean }>)
      .detail ?? { value: undefined, noProof: false };
    const target = this.#bridge?.activeEditor() ?? this.editor;
    if (!value || !target) return;
    if (target.state.selection.empty) return;
    target
      .chain()
      .setMark("textStyle", { language: { value }, noProof: noProof ? true : null })
      .run();
    this.#bridge?.focus();
    this.#syncStatusLanguage();
  };

  // ── Phonetic guide (拼音指南) ──

  /** The selection's phonetic state for the dialog: the per-character text,
   *  the readings already on its runs (blank where unannotated), the first
   *  ruby mark's alignment, and the selection bounds. Null when the selection
   *  is empty, spans paragraphs, or holds anything but text (the guide splits
   *  the run per character — mixed content and cross-paragraph ranges don't
   *  split). */
  #selectionPhonetic(): {
    chars: string[];
    readings: string[];
    alignment: string | null;
    from: number;
    to: number;
  } | null {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return null;
    const { from, to, empty, $from, $to } = editor.state.selection;
    if (empty || !$from.sameParent($to)) return null;
    const { doc } = editor.state;
    let plain = true;
    doc.nodesBetween(from, to, (node) => {
      // (nodesBetween yields the ancestors too — only an inline non-text node
      // inside the range blocks the split.)
      if (node.isInline && !node.isText) plain = false;
    });
    if (!plain) return null;
    const chars = doc.textBetween(from, to).split("");
    const readings = chars.map(() => "");
    let alignment: string | null = null;
    doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return;
      const ruby = (node.marks ?? []).find((m) => m.type.name === "ruby");
      if (!ruby) return;
      alignment ??= (ruby.attrs.alignment as string) ?? null;
      // This editor writes one node per base character carrying its whole
      // reading; a parsed multi-character node has no reliable per-character
      // split, so its reading lands whole on the first character.
      const start = Math.max(from, pos);
      const end = Math.min(to, pos + node.nodeSize);
      if (end > start && start - from < readings.length)
        readings[start - from] = String(ruby.attrs.text ?? "");
    });
    return { chars, readings, alignment, from, to };
  }

  /** Home → Font → Phonetic guide — open the per-character reading dialog
   *  (Word grays the button on an empty selection; a non-text or
   *  cross-paragraph selection is a no-op here). */
  readonly #onPhoneticOpen = (): void => {
    const dialog = this.shadowRoot?.querySelector("docen-phonetic-dialog") as unknown as {
      show(chars: string[], readings: string[], alignment: string | null): void;
    } | null;
    const state = this.#selectionPhonetic();
    if (!dialog || !state) return;
    dialog.show(state.chars, state.readings, state.alignment);
  };

  /** Phonetic dialog 确定 — split the selection into per-character runs, each
   *  carrying a ruby mark with its reading (a blank reading leaves that
   *  character unannotated). The base run's own marks ride every character;
   *  the annotation font is half the base size (Word's default). */
  readonly #onPhoneticOk = (event: Event): void => {
    const { chars, readings, alignment } =
      (
        event as CustomEvent<{
          chars?: string[];
          readings?: string[];
          alignment?: string;
        }>
      ).detail ?? {};
    const target = this.#bridge?.activeEditor() ?? this.editor;
    if (!target || !chars || !readings || chars.length === 0) return;
    const { from, to, empty, $from } = target.state.selection;
    if (empty) return;
    const carried = $from.marks();
    const baseSize =
      (carried.find((m) => m.type.name === "textStyle")?.attrs.size as number | null) ?? null;
    const { schema, tr } = target.state;
    const nodes = chars.map((ch, i) => {
      const marks = readings[i]
        ? [
            ...carried,
            schema.mark("ruby", {
              text: readings[i],
              alignment: alignment ?? "center",
              fontSize: baseSize != null ? Math.round(baseSize / 2) : null,
              baseFontSize: baseSize,
              raise: null,
              languageId: null,
              dirty: null,
            }),
          ]
        : carried;
      return schema.text(ch, marks);
    });
    const next = tr.replaceWith(from, to, nodes);
    next.setSelection(TextSelection.create(next.doc, from, from + chars.length));
    target.view.dispatch(next);
    this.#bridge?.focus();
  };

  /** Phonetic dialog 清除读音 — strip the ruby marks off the selection. */
  readonly #onPhoneticClear = (): void => {
    const target = this.#bridge?.activeEditor() ?? this.editor;
    if (!target || target.state.selection.empty) return;
    target.chain().unsetMark("ruby").run();
    this.#bridge?.focus();
  };

  /** Mirror the caret's proofing language into the status bar (Word shows the
   *  selection's language there). */
  #syncStatusLanguage(): void {
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.setAttribute("language", proofingLanguageName(this.#caretLanguage().value));
  }

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
      // readAsDataURL always yields a string — the guard narrows the union.
      if (typeof reader.result !== "string") return;
      const src = reader.result;
      // Natural size → attrs, clamped to the content width (Word inserts at
      // natural size but never wider than the frame, keeping the aspect).
      // Without explicit dimensions renderDocx falls back to a flat 400×300,
      // which distorts every non-default-shaped picture.
      const img = new Image();
      img.onload = (): void => {
        this.#bridge?.focus();
        const contentW = this.#flow?.contentWidthPx ?? 620;
        const scale = Math.min(1, contentW / Math.max(1, img.naturalWidth));
        this.editor?.commands.insertContent({
          type: "image",
          attrs: {
            src,
            width: Math.round(img.naturalWidth * scale),
            height: Math.round(img.naturalHeight * scale),
          },
        });
      };
      img.src = src;
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
    // Printing always outputs the paginated Print Layout pages (Word prints
    // the paper document whatever the view) — a continuous view re-projects
    // into print shape for the snapshot, then falls back.
    const mode = this.#viewMode();
    if (mode !== "print") {
      this.#stage?.setViewMode("print");
      this.#renderDoc(this.getJSON());
    }
    const shots = this.#stage?.printSnapshots() ?? [];
    if (mode !== "print") {
      this.#stage?.setViewMode(mode);
      this.#renderDoc(this.getJSON());
    }
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
   *  bytes can be awaited. While loading, an "Opening <name>" veil covers the
   *  canvas (Office shows the same message for a slow open) and the scroller
   *  stays frozen until the document is ready. */
  async openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void> {
    const name = input instanceof File ? input.name : undefined;
    this.#setProgress(t("status.opening", this).replace("{name}", name ?? "DOCX"));
    try {
      const buffer = input instanceof File ? await input.arrayBuffer() : input;
      // parseDOCX blocks the main thread — yield two frames so the veil paints
      // before the freeze (the bar's sweep is compositor-driven and keeps
      // moving through it).
      await this.#nextFrame();
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
      this.#applyOpenedJSON(parseMarkdown(text), name);
      await this.#nextFrame();
      this.#setProgress();
    } catch (err) {
      this.#setProgress();
      throw err;
    }
  }

  /** Open progress on the canvas veil — a label + indeterminate Fluent
   *  progress bar centered over the document area (Word centers its opening
   *  spinner the same way). Byte reads are a sliver of the load and parse/
   *  layout report nothing, so the bar never fakes a percentage. Clearing
   *  hides the veil. */
  #setProgress(label?: string): void {
    const root = this.shadowRoot;
    const veil = root?.querySelector<HTMLElement>(".load-veil");
    if (!veil || !root) return;
    if (label == null) {
      veil.hidden = true;
      return;
    }
    const labelEl = root.querySelector<HTMLElement>(".load-veil .load-label");
    if (!labelEl) return;
    veil.hidden = false;
    labelEl.textContent = label;
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
    // Panes are looked up by part, not position — several panes share the end
    // rail (properties + comments + clipboard + spelling).
    const part =
      id === "navigation"
        ? "nav-pane"
        : id === "comments"
          ? "comments-pane"
          : id === "clipboard"
            ? "clipboard-pane"
            : id === "proofing"
              ? "proofing-pane"
              : "props-pane";
    return this.shadowRoot?.querySelector(`docen-task-pane[part="${part}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
  }

  /** Apply a visibility state and dispatch `docen:taskpane-visibility-change`
   *  when it flips. The detail carries `visibilityMode: "taskpane"|"hidden"` to
   *  mirror `Office.VisibilityMode`. Idempotent — no event when state is
   *  unchanged. Opening a pane dismisses its rail-mates (Word's task panes are
   *  mutually exclusive per side — Comments replaces Properties, never
   *  stacks with it). */
  #setTaskpane(id: TaskPaneId, open: boolean): void {
    const pane = this.#paneEl(id);
    if (!pane || pane.open === open) return;
    if (open) {
      const side = pane.getAttribute("position") ?? "start";
      for (const other of this.shadowRoot?.querySelectorAll("docen-task-pane[open]") ?? []) {
        if (other !== pane && (other.getAttribute("position") ?? "start") === side) {
          (other as HTMLElement & { open: boolean }).open = false;
        }
      }
    }
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
   *  truth; the stage paints the ↵/→/· marks and break rows from it. */
  setShowMarks(on: boolean): void {
    if (this.hasAttribute("show-marks") === on) return;
    this.toggleAttribute("show-marks", on);
    this.#stage?.setShowMarks(on);
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
