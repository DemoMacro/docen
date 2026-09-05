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
  docxExtensions,
  effectiveRunProps,
  generateDOCX,
  generateMarkdown,
  normalizeDocument,
  parseDOCX,
  parseMarkdown,
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
import { attr, customElement } from "@microsoft/fast-element";
import type { Mark } from "@tiptap/pm/model";
import { EditorState, NodeSelection, TextSelection, type Transaction } from "@tiptap/pm/state";

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
import type { DrawingPropertiesState } from "../ui/components/workspace/drawing-properties-dialog";
import type { FontDialogPatch } from "../ui/components/workspace/font-dialog";
import { proofingLanguageName } from "../ui/components/workspace/language-dialog";
import type { LinkValues } from "../ui/components/workspace/link-dialog";
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
import { deepEq, dirtyPagesOf } from "./canvas/page-eq";
import {
  CanvasStage,
  type CanvasStageSection,
  type LaidFurnitureSection,
  layFurnitureSections,
} from "./canvas/stage";
import { documentStyles, documentTemplate, escapeHtml } from "./chrome";
import { ClipboardCommands } from "./commands/clipboard";
import { CommentsCommands } from "./commands/comments";
import { DesignCommands } from "./commands/design";
import { DialogCommands } from "./commands/dialogs";
import { NavigationCommands } from "./commands/navigation";
import { ReferencesCommands } from "./commands/references";
import { RevisionsCommands } from "./commands/revisions";
import { SectionCommands } from "./commands/sections";
import { SpellingCommands } from "./commands/spelling";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { pagesToPdf } from "./export-pdf";
import { tableAncestry, WIRED_DISPATCH } from "./extensions/commands";
import { LOCAL_HANDLED, READONLY_LIVE, SAVE_FORMATS, detectOpenFormat } from "./file-formats";
import { mergeSectionProperties } from "./page-setup";
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

/** Split buttons whose face carries no command of its own — the handler only
 *  exists for the drop-down variants' values (Word's menu buttons; a face
 *  click opens the menu instead of emitting a valueless command). */
const FACE_ONLY_SPLITS: ReadonlySet<string> = new Set(["autofit", "columns"]);

/**
 * Task pane identifiers, mirroring the Office `<TaskpaneId>` concept. The host
 * ships two built-in panes: `navigation` (start/left) and `properties` (end/right).
 */
export type TaskPaneId =
  | "navigation"
  | "properties"
  | "comments"
  | "clipboard"
  | "proofing"
  | "revisions";

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
  /** References-tab commands (citations/bibliography/index marking), split
   *  out of this class — see commands/references.ts. */
  readonly #spelling = new SpellingCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
  });
  readonly #navigation = new NavigationCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
    setTextSelection: (from, to) => this.#setTextSelection(from, to),
  });
  readonly #design = new DesignCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
  });
  readonly #comments = new CommentsCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
    showTaskpane: (id) => this.showTaskpane(id),
  });
  readonly #revisions = new RevisionsCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
  });
  readonly #references = new ReferencesCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
  });
  /** Dialog-commit commands (paragraph/font/table/Chinese layout/caption/
   *  cross-reference), split out of this class — see commands/dialogs.ts. */
  readonly #dialogs = new DialogCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
    syncStatusLanguage: () => this.#syncStatusLanguage(),
  });
  /** "This section" commands (sectPr read/write, page setup presets, the
   *  page-setup/columns/borders dialogs), split out of this class — see
   *  commands/sections.ts. */
  readonly #sections = new SectionCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
    flow: () => this.#flow,
  });
  /** Paste lanes, the paste-options bar, and the Office Clipboard pane,
   *  split out of this class — see commands/clipboard.ts. */
  readonly #clipboard = new ClipboardCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this,
  });
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
  /** Cached doc nodeSize + Office-style word count so caret-move transactions
   *  don't re-walk the whole document (recomputed only when content changes). */
  #lastDocSize = -1;
  #lastWords = 0;
  #unobserveLang?: () => void;
  /** Watches the host's `lang` attribute and forwards it to the internal
   *  <docen-workspace> + notifies locale observers. MutationObserver because
   *  @attr `lang` clashes with HTMLElement.lang (TS2416); manual
   *  observedAttributes would break FASTElement's @attr dispatch. */
  #langObserver?: MutationObserver;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured formatting + the listeners that apply it. Marks
  // and paragraph attrs capture together; sticky mode (double click) paints
  // every following selection until Esc or another painter click.
  #painterMarks: readonly Mark[] | null = null;
  #painterPara: Record<string, unknown> | null = null;
  #painterOff?: () => void;
  #painterKeyOff?: () => void;
  #painterSticky = false;
  #painterClickAt = 0;
  /** Current zoom level (percent) applied by the page stage's slot sizing. */
  #zoom = 100;
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
    // Alt+= inserts a blank inline equation (Word's Insert Equation shortcut).
    if (event.altKey && !event.ctrlKey && !event.metaKey && event.key === "=") {
      event.preventDefault();
      this.#insertEquation("plain");
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) return;
    // Ctrl+Shift+8 toggles formatting marks (Word). Shift+8 turns the key
    // into "*" on US layouts, so both spellings count.
    if (event.shiftKey && (event.key === "8" || event.key === "*")) {
      event.preventDefault();
      this.setShowMarks(!this.getShowMarks());
      return;
    }
    // Ctrl+F opens Find, Ctrl+H opens Find & Replace (Word behavior).
    if (event.key === "f" || event.key === "F") {
      event.preventDefault();
      this.#navigation.openSearch();
      return;
    }
    if (event.key === "h" || event.key === "H") {
      event.preventDefault();
      this.#navigation.openFindReplace();
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
      if (!this.#emitCancelable("docen:print")) void this.#print();
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

  /** Format Painter: a click captures the selection's run marks + paragraph
   *  formatting and arms a one-shot pointerup; the next non-empty selection
   *  receives both and disarms. A double click arms sticky mode — every
   *  following selection paints until Esc or another painter click (Word's
   *  format painter). A click while armed cancels. */
  #toggleFormatPainter(): void {
    const now = performance.now();
    const rapid = now - this.#painterClickAt < 500;
    this.#painterClickAt = now;
    if (this.#painterMarks) {
      if (rapid) {
        // Second click of a double click: stay armed, paint repeatedly.
        this.#painterSticky = true;
        return;
      }
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
    const $from = editor.state.selection.$from;
    if ($from.parent.type.name === "paragraph") {
      // Paragraph formatting paints too (Word: alignment, indent, spacing,
      // lists). The section-close markers belong to the target's position,
      // not the source's — stripped here, preserved on apply.
      const para = { ...($from.parent.attrs as Record<string, unknown>) };
      delete para.sectionProperties;
      delete para.sectionHeaders;
      delete para.sectionFooters;
      this.#painterPara = para;
    }
    this.#painterSticky = rapid;
    this.toggleAttribute("format-painter", true);
    const onUp = (): void => {
      const ed = this.#bridge?.activeEditor() ?? this.editor;
      if (ed) this.#applyFormatPainter(ed);
      if (this.#painterSticky) {
        // Stay armed: re-arm for the next selection. A bare caret click
        // (empty selection) consumes this listener without painting, exactly
        // like Word's sticky painter ignoring navigation clicks.
        this.addEventListener("pointerup", onUp, { once: true });
        this.#painterOff = () => this.removeEventListener("pointerup", onUp);
      } else {
        this.#stopFormatPainter();
      }
    };
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") this.#stopFormatPainter();
    };
    this.addEventListener("pointerup", onUp, { once: true });
    this.addEventListener("keydown", onKey);
    this.#painterOff = () => this.removeEventListener("pointerup", onUp);
    this.#painterKeyOff = () => this.removeEventListener("keydown", onKey);
  }

  /** Stamp the captured marks + paragraph attrs onto the current selection. */
  #applyFormatPainter(ed: Editor): void {
    const { from, to, empty } = ed.state.selection;
    if (empty || (!this.#painterMarks && !this.#painterPara)) return;
    const tr = ed.state.tr;
    for (const mark of this.#painterMarks ?? []) tr.addMark(from, to, mark);
    if (this.#painterPara) {
      ed.state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type.name !== "paragraph") return;
        const next: Record<string, unknown> = { ...this.#painterPara! };
        // The target keeps its own section-close markers.
        const target = node.attrs as Record<string, unknown>;
        for (const key of ["sectionProperties", "sectionHeaders", "sectionFooters"]) {
          if (target[key] != null) next[key] = target[key];
        }
        tr.setNodeMarkup(pos, undefined, next);
      });
    }
    ed.view.dispatch(tr);
  }

  #stopFormatPainter(): void {
    this.#painterMarks = null;
    this.#painterPara = null;
    this.#painterSticky = false;
    this.removeAttribute("format-painter");
    this.#painterOff?.();
    this.#painterOff = undefined;
    this.#painterKeyOff?.();
    this.#painterKeyOff = undefined;
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
        this.#spelling.replace(event.detail)) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:ignore-all", () => this.#spelling.ignore("ignore"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:add", () => this.#spelling.ignore("add"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:nav", ((event: CustomEvent<number>) =>
        this.#spelling.goto(this.#spelling.activeIndex() + event.detail)) as EventListener);

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
      onOutlineUpdate: (anchors) => this.#navigation.renderOutline(anchors),
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
      onRichPaste: (source) => this.#clipboard.showPasteOptions(source),
      onClipboardCollect: (item) => this.#clipboard.collect(item),
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
      drawingBoxOf: (para, index, kind) => {
        const stage = this.#stage;
        const hit = stage?.drawingBoxOf(para, index, kind) ?? null;
        if (hit || !stage) return hit;
        // A write-back (a resize) re-laid the doc and re-objected every
        // paragraph — the caller's hit reference is stale. Re-resolve the
        // PM selection (still a NodeSelection on the same drawing) against
        // the fresh boxes so the frame follows the resized picture.
        const sel = this.editor?.state.selection;
        if (!(sel instanceof NodeSelection)) return null;
        return (
          stage
            .drawingBoxes()
            .find((box) => this.#drawingNodePos(box.para, box.index, box.kind) === sel.from) ?? null
        );
      },
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
      this.#navigation.onOutlineSelect as EventListener,
    );
    // Nav-pane search → Find (live highlight, next/prev, results count).
    this.addEventListener("navigation:search", this.#navigation.onSearch as EventListener);
    this.addEventListener("navigation:find", this.#navigation.onFind as EventListener);
    // Comments pane → create/cancel (compose box), select (scroll to range),
    // update/delete (inline card actions), reply/resolve (thread actions).
    this.addEventListener("comment:create", this.#comments.onCommentCreate as EventListener);
    this.addEventListener("comment:cancel", this.#comments.onCommentCancel as EventListener);
    this.addEventListener("comment:select", this.#comments.onCommentSelect as EventListener);
    this.addEventListener("comment:update", this.#comments.onCommentUpdate as EventListener);
    this.addEventListener("comment:delete", this.#comments.onCommentDelete as EventListener);
    this.addEventListener("comment:reply", this.#comments.onCommentReply as EventListener);
    this.addEventListener("comment:resolve", this.#comments.onCommentResolve as EventListener);
    // Reviewing pane → select (scroll to the revision), accept/reject (the
    // by-id command form — one transaction per card action).
    this.addEventListener("revision:select", this.#revisions.onRevisionSelect as EventListener);
    this.addEventListener("revision:accept", this.#revisions.onRevisionAccept as EventListener);
    this.addEventListener("revision:reject", this.#revisions.onRevisionReject as EventListener);
    // Office Clipboard pane → paste one entry / paste all / clear.
    this.addEventListener("clipboard:paste", this.#clipboard.onPanePaste as EventListener);
    this.addEventListener("clipboard:paste-all", this.#clipboard.onPanePasteAll as EventListener);
    this.addEventListener("clipboard:clear", this.#clipboard.onPaneClear as EventListener);
    // Click a Results entry → jump to that match (delegated on the container).
    this.shadowRoot!.querySelector(".search-results")?.addEventListener(
      "click",
      this.#navigation.onSearchResultClick as EventListener,
    );
    // Find & Replace dialog → Replace / Replace All (prosemirror-search).
    this.shadowRoot!.querySelector("docen-find-replace-dialog")?.addEventListener(
      "find-replace:action",
      this.#navigation.onFindReplace as EventListener,
    );
    // Options dialog — ok (UI language + theme).
    this.shadowRoot!.querySelector("docen-options-dialog")?.addEventListener(
      "options:ok",
      this.#onOptionsOk as EventListener,
    );
    // Document Inspector (检查问题) — the findings dialog's removal buttons.
    this.shadowRoot!.querySelector("docen-inspect-dialog")?.addEventListener(
      "inspect:clear-comments",
      this.#onInspect as EventListener,
    );
    this.shadowRoot!.querySelector("docen-inspect-dialog")?.addEventListener(
      "inspect:accept-revisions",
      this.#onInspect as EventListener,
    );
    // Language dialog — commit the selection's proofing language (w:lang).
    this.shadowRoot!.querySelector("docen-language-dialog")?.addEventListener(
      "language:ok",
      this.#dialogs.onLanguageOk as EventListener,
    );
    // Phonetic guide dialog — split the selection into per-character ruby
    // runs, or strip the guides off it.
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:ok",
      this.#dialogs.onPhoneticOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:clear",
      this.#dialogs.onPhoneticClear as EventListener,
    );
    // Two Lines in One dialog — pack the selection's text into two half-size
    // lines (双行合一 / 合并字符).
    this.shadowRoot!.querySelector("docen-two-in-one-dialog")?.addEventListener(
      "two-in-one:ok",
      this.#dialogs.onTwoInOneOk as EventListener,
    );
    // Define New Multilevel List dialog — register the levels as a document
    // numbering definition and stamp the selection with it.
    this.shadowRoot!.querySelector("docen-define-list-dialog")?.addEventListener(
      "define-list:ok",
      this.#dialogs.onDefineListOk as EventListener,
    );
    // Caption dialog — seed a Caption-styled paragraph with a SEQ field next
    // to the caret's paragraph.
    this.shadowRoot!.querySelector("docen-caption-dialog")?.addEventListener(
      "caption:ok",
      this.#dialogs.onCaptionOk as EventListener,
    );
    // Cross-reference dialog — seed a cached REF/PAGEREF field at the caret.
    this.shadowRoot!.querySelector("docen-cross-reference-dialog")?.addEventListener(
      "cross-ref:ok",
      this.#dialogs.onCrossRefOk as EventListener,
    );
    // Sources dialog — write the bibliography source list (attrs.bibliography)
    // and seed a cached CITATION field at the caret.
    this.shadowRoot!.querySelector("docen-sources-dialog")?.addEventListener(
      "sources:ok",
      this.#references.onSourcesOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-sources-dialog")?.addEventListener(
      "citation:ok",
      this.#references.onCitationOk as EventListener,
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
      this.#dialogs.onParagraphOk as EventListener,
    );
    // Page Setup dialog — write the committed geometry into the current
    // section (the Custom Margins / More Paper Sizes entries open it).
    this.shadowRoot!.querySelector("docen-page-setup-dialog")?.addEventListener(
      "page-setup:ok",
      this.#sections.onPageSetupOk as EventListener,
    );
    // Table grid — insert the picked shape through insert-table.
    this.shadowRoot!.querySelector("docen-table-dialog")?.addEventListener(
      "table-grid:insert",
      this.#onTableInsert as EventListener,
    );
    // Columns dialog — write the committed layout into the current section.
    this.shadowRoot!.querySelector("docen-columns-dialog")?.addEventListener(
      "columns:ok",
      this.#sections.onColumnsOk as EventListener,
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
      this.#dialogs.onFontOk as EventListener,
    );
    // Table Properties dialog — rewrite the caret table's alignment/indent.
    this.shadowRoot!.querySelector("docen-table-properties-dialog")?.addEventListener(
      "table-properties:ok",
      this.#dialogs.onTablePropertiesOk as EventListener,
    );
    // Size-and-Position dialog — restamp the selected drawing's geometry.
    this.shadowRoot!.querySelector("docen-drawing-properties-dialog")?.addEventListener(
      "drawing-properties:ok",
      this.#dialogs.onDrawingPropertiesOk as EventListener,
    );
    // Borders and Shading dialog — stamp the border/page/shading tab.
    this.shadowRoot!.querySelector("docen-borders-shading-dialog")?.addEventListener(
      "borders-shading:ok",
      this.#sections.onBordersShadingOk as EventListener,
    );
    // Custom watermark dialog — clear/stamp the header watermark.
    this.shadowRoot!.querySelector("docen-watermark-dialog")?.addEventListener(
      "watermark:ok",
      this.#design.onWatermarkOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "zoom:open",
      this.#onZoomOpen as EventListener,
    );
    // Status-bar word count → the statistics dialog (Word).
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "wordcount:open",
      this.#onWordCountOpen as EventListener,
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
    this.editor?.on("selectionUpdate", this.#comments.syncActiveCommentCard);
    this.editor?.on("selectionUpdate", this.#revisions.syncActiveRevision);
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
    const targetPos = this.#sections.sectionSectPrPos();
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
    const targetPos = this.#sections.sectionSectPrPos();
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

  /** The selected drawing's geometry in the dialog's display unit (cm for
   *  size and offsets, degrees for rotation) — an image sizes in px attrs
   *  while a shape payload sizes in EMU; null on any other selection. */
  #drawingStateOf(): DrawingPropertiesState | null {
    const sel = this.editor?.state.selection;
    if (!(sel instanceof NodeSelection)) return null;
    const attrs = sel.node.attrs as Record<string, unknown>;
    const shape = attrs.wpsShape as Record<string, unknown> | null | undefined;
    const floating = (shape ? shape.floating : attrs.floating) as Record<string, unknown> | null;
    if (!floating) return null;
    const EMU_PER_CM = 360000;
    const PX_PER_CM = 96 / 2.54;
    const sizeDiv = shape ? EMU_PER_CM : PX_PER_CM;
    const t = (shape?.transformation ?? {}) as Record<string, unknown>;
    const num = (v: unknown): number => (typeof v === "number" && Number.isFinite(v) ? v : 0);
    const cm = (v: number): number => Math.round((v / sizeDiv) * 100) / 100;
    const offsetCm = (v: unknown): number => Math.round((num(v) / EMU_PER_CM) * 100) / 100;
    const hPos = floating.horizontalPosition as Record<string, unknown> | undefined;
    const vPos = floating.verticalPosition as Record<string, unknown> | undefined;
    return {
      widthCm: cm(shape ? num(t.width) : num(attrs.width)),
      heightCm: cm(shape ? num(t.height) : num(attrs.height)),
      rotationDeg: num(shape ? t.rotation : attrs.rotation),
      offsetHCm: offsetCm(hPos?.offset),
      offsetVCm: offsetCm(vPos?.offset),
    };
  }

  /** The selected node when it's a source-carrying image (crop's target) —
   *  null for shapes and every other selection. */
  #selectedImage(): { src: string } | null {
    const sel = this.editor?.state.selection;
    if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return null;
    const src = sel.node.attrs.src;
    return typeof src === "string" && src ? { src } : null;
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
    this.#comments.syncCommentsPane();
    this.#revisions.syncRevisionsPane();
    this.#spelling.schedule();
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
    this.#clipboard.hidePasteOptions();
    this.#langObserver?.disconnect();
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.shadowRoot?.removeEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot?.removeEventListener("change", this.#onChange as EventListener);
    this.#fileInput?.removeEventListener("change", this.#onFileChange);
    this.#imageInput?.removeEventListener("change", this.#onImageChange);
    this.shadowRoot
      ?.querySelector("docen-outline")
      ?.removeEventListener("outline:select", this.#navigation.onOutlineSelect as EventListener);
    this.removeEventListener("navigation:search", this.#navigation.onSearch as EventListener);
    this.removeEventListener("navigation:find", this.#navigation.onFind as EventListener);
    this.shadowRoot
      ?.querySelector(".search-results")
      ?.removeEventListener("click", this.#navigation.onSearchResultClick as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#navigation.onFindReplace as EventListener);
    this.shadowRoot
      ?.querySelector("docen-options-dialog")
      ?.removeEventListener("options:ok", this.#onOptionsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("language:open", this.#onLanguageOpen as EventListener);
    this.shadowRoot
      ?.querySelector("docen-language-dialog")
      ?.removeEventListener("language:ok", this.#dialogs.onLanguageOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:ok", this.#dialogs.onPhoneticOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:clear", this.#dialogs.onPhoneticClear as EventListener);
    this.shadowRoot
      ?.querySelector("docen-two-in-one-dialog")
      ?.removeEventListener("two-in-one:ok", this.#dialogs.onTwoInOneOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-define-list-dialog")
      ?.removeEventListener("define-list:ok", this.#dialogs.onDefineListOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-caption-dialog")
      ?.removeEventListener("caption:ok", this.#dialogs.onCaptionOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-cross-reference-dialog")
      ?.removeEventListener("cross-ref:ok", this.#dialogs.onCrossRefOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-sources-dialog")
      ?.removeEventListener("sources:ok", this.#references.onSourcesOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-sources-dialog")
      ?.removeEventListener("citation:ok", this.#references.onCitationOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-symbol-dialog")
      ?.removeEventListener("symbol:insert", this.#onSymbolInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paragraph-dialog")
      ?.removeEventListener("paragraph:ok", this.#dialogs.onParagraphOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paste-special-dialog")
      ?.removeEventListener("paste-special:ok", this.#onPasteSpecialOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-font-dialog")
      ?.removeEventListener("font:ok", this.#dialogs.onFontOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-properties-dialog")
      ?.removeEventListener(
        "table-properties:ok",
        this.#dialogs.onTablePropertiesOk as EventListener,
      );
    this.shadowRoot
      ?.querySelector("docen-borders-shading-dialog")
      ?.removeEventListener(
        "borders-shading:ok",
        this.#sections.onBordersShadingOk as EventListener,
      );
    this.shadowRoot
      ?.querySelector("docen-watermark-dialog")
      ?.removeEventListener("watermark:ok", this.#design.onWatermarkOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-page-setup-dialog")
      ?.removeEventListener("page-setup:ok", this.#sections.onPageSetupOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-dialog")
      ?.removeEventListener("table-grid:insert", this.#onTableInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-columns-dialog")
      ?.removeEventListener("columns:ok", this.#sections.onColumnsOk as EventListener);
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
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("wordcount:open", this.#onWordCountOpen as EventListener);
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
    this.editor?.off("selectionUpdate", this.#comments.syncActiveCommentCard);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.shadowRoot
      ?.querySelector("docen-ribbon")
      ?.removeEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    this.#fontSyncCleanup?.();
    this.#fontSyncCleanup = undefined;
    this.#stopFormatPainter();
    this.#navigation.dispose();
    this.#spelling.dispose();
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
              <fluent-menu-button
                slot="trigger"
                appearance="subtle"
                style="max-width:36vw;overflow:hidden;white-space:nowrap"
                title="${escapeHtml(filename)}"
              >${escapeHtml(filename)}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new">${t("header.new", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="open">${t("header.open", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="save-as">${t("header.save-as", this)}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-markdown">${t("header.save-as-markdown", this)}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-pdf">${t("header.save-as-pdf", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="print">${t("header.print", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="share">${t("header.share", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="properties">${t("header.properties", this)}</fluent-menu-item>
                <fluent-menu-item data-event="inspect-document">${t("header.inspect", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="options">${t("header.options", this)}</fluent-menu-item>
                <fluent-menu-item data-event="close">${t("header.close", this)}</fluent-menu-item>
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
    root
      .querySelector('docen-task-pane[part="revisions-pane"]')
      ?.setAttribute("title", t("revisions.title", this));
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
    const attrs = this.editor?.state.doc.attrs as
      | {
          sectionProperties?: { titlePage?: boolean };
          documentExtras?: { settings?: { evenAndOddHeaders?: boolean } };
        }
      | undefined;
    // titlePage is a sectPr flag; evenAndOddHeaders lives in settings.xml
    // (toggled through documentExtras — see SectionsCommands.toggleSectionFlag).
    const sp = attrs?.sectionProperties;
    const oddEven = attrs?.documentExtras?.settings?.evenAndOddHeaders;
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
            checked: !!oddEven,
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
      oddEvenCb?.toggleAttribute("checked", !!oddEven);
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
    name:
      | "docen:save"
      | "docen:save-as"
      | "docen:open"
      | "docen:new"
      | "docen:print"
      | "docen:close",
    detail?: { format?: "docx" | "markdown" | "pdf" },
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

  /** Toggle a task pane open/closed (ribbon View → toggle-navigation). */
  #togglePane(id: TaskPaneId): void {
    this.#setTaskpane(id, !this.getTaskpaneState(id));
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

  readonly #onWordCountOpen = (): void => {
    this.#showWordCount();
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
    void this.#clipboard.paste(event.detail === "text");
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
    const id = this.#dialogs.nextBookmarkId(editor);
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
      // Alt+= inserts a blank equation — one empty run to type into.
      plain: { text: "" },
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
    // Word: a right-click on a picture selects it and shows the picture menu
    // instead of the text menu — the caret-based branches below never run.
    if (this.#bridge?.selectDrawingAtClient(event.clientX, event.clientY)) {
      const items: RibbonMenuItem[] = [];
      // The clipboard entries work on a NodeSelection: the slice payload
      // carries the whole drawing, deleteSelection cuts it.
      items.push({ text: t("context.cut", this), event: "cut" });
      items.push({ text: t("context.copy", this), event: "copy" });
      items.push({ text: "-" });
      items.push({ text: t("context.bring-forward", this), event: "bring-forward" });
      items.push({ text: t("context.send-backward", this), event: "send-backward" });
      // The numeric layout dialog needs an offset-anchored floating drawing —
      // an inline picture has no offset to edit.
      items.push({
        text: t("context.size-position", this),
        event: "drawing-properties",
        ...(this.#drawingStateOf() ? {} : { disabled: true }),
      });
      // Crop edits the picture's source — a shape has no source of its own.
      items.push({
        text: t("context.crop", this),
        event: "drawing-crop",
        ...(this.#selectedImage() ? {} : { disabled: true }),
      });
      items.push({ text: "-" });
      items.push({ text: t("context.delete-picture", this), event: "delete-picture" });
      menu.setAttribute("items", JSON.stringify(items));
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
      else this.#navigation.openSearch();
      return;
    }
    // Replace (ribbon Home → Editing → Replace, or Ctrl+H) → Find & Replace dialog.
    if (name === "replace") {
      this.#navigation.openFindReplace();
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
      if (value === "more") this.#sections.openPageSetup();
      else this.#sections.setPageSize(value);
      return;
    }
    if (name === "orientation") {
      this.#sections.setOrientation(value);
      return;
    }
    if (name === "margins") {
      if (value === "custom") this.#sections.openPageSetup();
      else this.#sections.setMargins(value);
      return;
    }
    // Columns presets (the Layout tab's Columns menu: one/two/three);
    // More Columns opens the dialog prefilled from the current section.
    if (name === "columns") {
      if (value === "more") this.#sections.openColumnsDialog();
      else {
        const count = Number(value);
        if (count >= 1 && count <= 9) this.#sections.setColumnCount(count);
      }
      return;
    }
    // Line Numbers toggle (the Layout tab's Line Numbers button).
    if (name === "line-numbers") {
      this.#sections.toggleLineNumbers();
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
    // Table of Figures — insert/update the caption directory (the TOC field's
    // \c switch): same story routing and post-insert update pass as the TOC.
    if (name === "table-of-figures" || name === "update-figures") {
      const target = this.#bridge?.activeEditor() ?? editor;
      const pageOf = (pos: number): number | null => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const tabPositionTw = this.#flow
        ? Math.round(this.#flow.contentWidthPx / twipToPx(1))
        : undefined;
      const ran = target.commands[name](pageOf, tabPositionTw);
      if (name === "table-of-figures" && ran) {
        requestAnimationFrame(() =>
          requestAnimationFrame(() => target.commands["update-figures"](pageOf, tabPositionTw)),
        );
      }
      return;
    }
    // Index — Mark Entry prompts for the entry text and seeds an XE field at
    // the selection (the invisible marker Word hides from the page); insert
    // and update collect the XE fields into the Index-styled entry block.
    if (name === "mark-entry" || name === "insert-index" || name === "update-index") {
      const target = this.#bridge?.activeEditor() ?? editor;
      if (name === "mark-entry") {
        this.#references.markIndexEntry(target);
        return;
      }
      const pageOf = (pos: number): number | null => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const tabPositionTw = this.#flow
        ? Math.round(this.#flow.contentWidthPx / twipToPx(1))
        : undefined;
      const ran = target.commands[name](pageOf, tabPositionTw);
      if (!ran) window.alert(t("index.empty", this));
      return;
    }
    // Header/Footer — the split's main action opens the story on the caret's
    // page; the drop-down carries remove + the slot-visibility flags.
    if (name === "header" || name === "footer") {
      if (value === "title-page" || value === "odd-even") {
        this.#sections.toggleSectionFlag(
          value === "title-page" ? "titlePage" : "evenAndOddHeaders",
        );
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
        this.#sections.toggleSectionFlag(
          value === "title-page" ? "titlePage" : "evenAndOddHeaders",
        );
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
    // Size and Position — open the drawing dialog prefilled from the selected
    // floating drawing; the commit arrives via drawing-properties:ok
    // (drawing-properties-apply).
    if (name === "drawing-properties") {
      const dialog = this.shadowRoot?.querySelector("docen-drawing-properties-dialog") as {
        show(state: DrawingPropertiesState): void;
      } | null;
      const state = this.#drawingStateOf();
      if (dialog && state) dialog.show(state);
      return;
    }
    // Crop — the bridge enters crop mode on the selected image (the overlay
    // previews the full source; Enter / a press outside commits, Esc cancels).
    if (name === "drawing-crop") {
      this.#bridge?.enterCropMode();
      return;
    }
    // Bookmark — prompt for a name and wrap the selection with a
    // bookmarkStart/bookmarkEnd pair (Word's Insert → Bookmark).
    if (name === "bookmark") {
      this.#insertBookmark();
      return;
    }
    // Caption — open the dialog; the commit arrives via caption:ok
    // (#dialogs.onCaptionOk, References → Captions → Insert Caption).
    if (name === "insert-caption") {
      (this.shadowRoot?.querySelector("docen-caption-dialog") as { show(): void } | null)?.show();
      return;
    }
    // Cross-reference — open the dialog over the document's bookmarks; the
    // commit arrives via cross-ref:ok (#dialogs.onCrossRefOk).
    if (name === "cross-reference") {
      (
        this.shadowRoot?.querySelector("docen-cross-reference-dialog") as {
          show(targets: { name: string; text: string; kind: string }[]): void;
        } | null
      )?.show(this.#dialogs.crossReferenceTargets());
      return;
    }
    // Source Manager / Insert Citation — the same dialog in two modes (Word's
    // References → Citations & Bibliography group); commits arrive via
    // sources:ok / citation:ok (#references.onSourcesOk / onCitationOk).
    if (name === "manage-sources" || name === "insert-citation") {
      (
        this.shadowRoot?.querySelector("docen-sources-dialog") as {
          show(mode: "manage" | "cite", sources: unknown[]): void;
        } | null
      )?.show(
        name === "insert-citation" ? "cite" : "manage",
        this.#references.bibliographySources(),
      );
      return;
    }
    // Bibliography — insert (or rebuild) the Bibliography-styled block beside
    // the caret from the document's sources (#references.insertBibliography).
    if (name === "bibliography") {
      this.#references.insertBibliography();
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
      this.#design.setPageColor(
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
      this.#sections.openBordersDialog(name === "page-border" ? "page" : "border");
      return;
    }
    if (name === "page-border") {
      this.#sections.setPageBorders(value);
      return;
    }
    // Paragraph Spacing presets — stamp the styles' docDefaults paragraph
    // spacing (Word's Design → Paragraph Spacing; the document-level default
    // every paragraph without explicit spacing inherits).
    if (name === "paragraph-spacing") {
      this.#design.setParagraphSpacing(typeof value === "string" ? value : undefined);
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
        this.#design.openWatermarkDialog();
        return;
      }
      this.#design.setWatermark(typeof value === "string" ? value : undefined);
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
      this.#comments.insertComment();
      return;
    }
    // The trailing title-bar "comment" button toggles the comments pane
    // (Word's sidebar) — it lists every comment, it does not create one.
    if (name === "comment") {
      this.#togglePane("comments");
      return;
    }
    if (name === "edit-comment") {
      this.#comments.editComment();
      return;
    }
    if (name === "delete-comment") {
      this.#comments.deleteComment();
      return;
    }
    if (name === "previous-comment") {
      this.#comments.jumpComment("previous");
      return;
    }
    if (name === "next-comment") {
      this.#comments.jumpComment("next");
      return;
    }
    // Review → Show Comments: toggle the comments pane (Word's sidebar).
    if (name === "show-comments") {
      this.#setTaskpane("comments", !this.getTaskpaneState("comments"));
      return;
    }
    // Review → Reviewing Pane: toggle the revisions pane (Word's vertical
    // reviewing pane listing every tracked change).
    if (name === "reviewing-pane") {
      this.#togglePane("revisions");
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
      void this.#clipboard.paste(value === "keep-text-only");
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
      this.#spelling.run();
      this.#setTaskpane("proofing", true);
      const from = this.editor?.state.selection.from ?? 0;
      const issues = this.#spelling.issues();
      if (issues.length) {
        const first = issues.find((issue) => issue.from >= from) ?? issues[0];
        this.#spelling.goto(issues.indexOf(first));
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
      this.#dialogs.phoneticOpen();
      return;
    }
    // Chinese Layout (中文版式, Home → Paragraph): the two-lines-in-one
    // dialog over the selection (合并字符 rides the same dialog).
    if (name === "two-lines-in-one") {
      this.#dialogs.twoInOneOpen();
      return;
    }
    // Multilevel List gallery (Home → Paragraph): the Define New Multilevel
    // List dialog — its last entry.
    if (name === "define-new-list") {
      this.#dialogs.defineListOpen();
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
      case "save-as-pdf":
        if (!this.#emitCancelable("docen:save-as", { format: "pdf" })) void this.#saveAsPdf();
        break;
      case "print":
        if (!this.#emitCancelable("docen:print")) void this.#print();
        break;
      case "properties":
        // Word's File → Info: the document properties pane.
        this.showTaskpane("properties");
        break;
      case "inspect-document":
        this.#inspectDocument();
        break;
      case "share":
        void this.#share();
        break;
      case "close":
        this.#closeDocument();
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
    await this.#saveBlob(data as BlobPart, cfg, true);
  }

  /** Write a finished blob out through the File System Access picker (adopting
   *  the picked name as the filename when `adoptName`), falling back to a
   *  plain download where the picker doesn't exist. `adoptName` is false for
   *  format exports (PDF) — saving a copy doesn't rename the document. */
  async #saveBlob(
    data: BlobPart,
    cfg: { description: string; mime: string; ext: string },
    adoptName: boolean,
  ): Promise<void> {
    const blob = new Blob([data], { type: cfg.mime });
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
        if (adoptName) {
          this.setAttribute("filename", handle.name);
          this.#renderChrome();
        }
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

  /** Export as PDF (filename menu → Export as PDF): the paginated print
   *  snapshots flatten into a PDF blob. Same view round-trip as #print — a
   *  non-print view re-projects into print shape for the snapshot, then falls
   *  back. The export never renames the document. */
  async #saveAsPdf(): Promise<void> {
    const mode = this.#viewMode();
    if (mode !== "print") {
      this.#stage?.setViewMode("print");
      this.#renderDoc(this.getJSON());
    }
    const shots = (await this.#stage?.printSnapshots()) ?? [];
    if (mode !== "print") {
      this.#stage?.setViewMode(mode);
      this.#renderDoc(this.getJSON());
    }
    if (shots.length === 0) return;
    const blob = await pagesToPdf(shots);
    await this.#saveBlob(blob, SAVE_FORMATS.pdf, false);
  }

  /** Filename menu → Share: the Web Share sheet where the platform has one
   *  (title only — the document body is not uploaded); otherwise copy the
   *  document URL (Word for the web's share = share a link). */
  async #share(): Promise<void> {
    const title = this.getAttribute("filename") ?? t("header.doc-name", this);
    const nav = navigator as Navigator & { share?: (data: ShareData) => Promise<void> };
    if (typeof nav.share === "function") {
      try {
        await nav.share({ title });
        return;
      } catch {
        return; // the user dismissed the sheet (AbortError)
      }
    }
    try {
      await navigator.clipboard.writeText(location.href);
    } catch {
      // Clipboard denied — nothing else to offer.
    }
  }

  /** Filename menu → Close: end the editing session. A host takes over via
   *  docen:close; otherwise the document resets to a blank slate (Word's Close
   *  closes the window — the browser element's equivalent). Unsaved work is
   *  confirmed away — there is no dirty-save model to offer. */
  #closeDocument(): void {
    if (this.#emitCancelable("docen:close")) return;
    if (this.#jsonDirty && !window.confirm(t("close.confirm", this))) return;
    this.setAttribute("filename", t("header.doc-name", this));
    this.#renderChrome();
    this.setJSON({ type: "doc", content: [{ type: "paragraph" }] });
  }

  /** The Document Inspector's scan: comment cards in documentExtras and the
   *  distinct revision records (w:ins/w:del ids) in the doc. */
  #inspectFindings(): { comments: number; revisions: number } {
    const comments = (
      (this.editor?.state.doc.attrs ?? {}) as {
        documentExtras?: { comments?: unknown[] };
      }
    ).documentExtras?.comments?.length;
    const ids = new Set<string>();
    this.editor?.state.doc.descendants((node) => {
      if (!node.isText) return true;
      for (const mark of node.marks) {
        const name = mark.type.name;
        if (name === "insertion" || name === "deletion") {
          ids.add(`${name}:${String((mark.attrs as { id?: unknown }).id)}`);
        }
      }
      return true;
    });
    return { comments: comments ?? 0, revisions: ids.size };
  }

  /** Filename menu → Inspect Document (Word's 检查问题): scan, then show the
   *  findings dialog; its removal buttons come back as events (#onInspect). */
  #inspectDocument(): void {
    const dialog = this.shadowRoot?.querySelector("docen-inspect-dialog") as unknown as {
      setAttribute(name: string, value: string): void;
      show(): void;
    } | null;
    if (!dialog) return;
    dialog.setAttribute("findings", JSON.stringify(this.#inspectFindings()));
    dialog.show();
  }

  /** The inspector dialog's removal buttons — clear comments / accept all
   *  revisions, then re-hand the scan so the counts fall to zero in place. */
  readonly #onInspect = (event: Event): void => {
    if (event.type === "inspect:clear-comments") {
      this.#comments.deleteAllComments();
    } else if (event.type === "inspect:accept-revisions") {
      const commands = this.editor?.commands;
      if (commands)
        (commands as unknown as Record<string, () => unknown>)["accept-all-changes"]?.();
    }
    const dialog = this.shadowRoot?.querySelector("docen-inspect-dialog");
    dialog?.setAttribute("findings", JSON.stringify(this.#inspectFindings()));
  };

  /** Print only the document pages — never the ribbon/chrome. Each page
   *  canvas rasterizes into a hidden print-only iframe (one image per page at
   *  the page's true paper size, @page margin 0), so the browser's print
   *  dialog receives exactly the paginated document, like Word's print
   *  output. */
  async #print(): Promise<void> {
    // Printing always outputs the paginated Print Layout pages (Word prints
    // the paper document whatever the view) — a continuous view re-projects
    // into print shape for the snapshot, then falls back.
    const mode = this.#viewMode();
    if (mode !== "print") {
      this.#stage?.setViewMode("print");
      this.#renderDoc(this.getJSON());
    }
    const shots = (await this.#stage?.printSnapshots()) ?? [];
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
              : id === "revisions"
                ? "revisions-pane"
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
