/**
 * `<docen-presentation>` — a turnkey PPTX editor web component.
 *
 * Wires the Fluent UI host (title-bar + ribbon + document-area + status-bar)
 * to the pptx route: parse through @docen/pptx, project the slides into
 * drawing members, and paint them with the shared core painter on a Leafer
 * stage. The title bar drives file I/O (open) and language switching, ribbon
 * commands route through the host, and unwired commands grey out honestly —
 * the same contract as `<docen-document>`. Editing (selection, gestures,
 * engine transactions) rides on this base in later batches.
 */

import {
  browserFontMetrics,
  familyOfSlot,
  leaferBaselinePadPx,
  type LayoutBlock,
  type LayoutDrawingMember,
} from "@docen/layout";
import {
  generatePresentation,
  parsePresentation,
  projectPresentation,
  tableGridOf,
  memberAt,
  memberByPath,
  type PresentationOptions,
  type ProjectedPresentation,
  type SlideOptions,
  type SlideChild,
  type TableOptions,
  type TableCellOptions,
  type TransitionOptions,
  type TransitionType,
} from "@docen/pptx";
import { customElement, observable } from "@microsoft/fast-element";
import type { DataType } from "@office-open/core";
import type { ShapeType, TextBodyOptions, TextRunOptions } from "@office-open/core/drawing";
import { App, type IGroup } from "leafer-ui";

import { renderRibbonFromSchema } from "../document/ribbon";
import type { Box } from "../drawing/geometry";
import { DrawingOverlay } from "../drawing/overlay";
import {
  AddinHost,
  applyTheme,
  mergeRibbonSchema,
  notifyLocaleChange,
  observeLang,
  registerComponents,
  resolveTheme,
  t,
} from "../ui";
import type { LinkValues } from "../ui/components/workspace/link-dialog";
import { escapeHtml, presentationStyles, presentationTemplate } from "./chrome";
import {
  bodyParagraphsOf,
  cellParagraphsOf,
  cellTextOf,
  childRunsOf,
  deleteSlideAt,
  duplicateSlideAt,
  firstCellRunSizeOf,
  firstRunSizeOf,
  insertSlideAt,
  makePicture,
  makeShape,
  makeTable,
  makeTextBox,
  reorderChild,
  runsIn,
  setLineSpacingPercent,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  shapeTextOf,
  slideNotesOf,
  toggleBullet,
  toggleNumbering,
  toggleRunFlag,
  toggleRunStyle,
  writeCellText,
  writeShapeText,
  writeSlideNotes,
} from "./commands";
import {
  captureGeometry,
  hitSlide,
  offsetChild,
  resizeChild,
  restoreGeometry,
  rotateChild,
  slideHits,
} from "./hit-test";
import { presentationRibbonTabs } from "./ribbon";
import {
  paintSlideDeck,
  repaintSlideAt,
  SLIDE_GAP_PX,
  THUMB_GAP_PX,
  THUMB_WIDTH_PX,
} from "./slides-panel";
// Side-effect: register the presentation translation tables.
import "./i18n";

const ZOOM_MIN = 10;
const ZOOM_MAX = 500;

/** The drawing grid pitch: PowerPoint's 0.5" grid in slide px (96/inch). */
const GRID_PITCH_PX = 48;

/** Undo steps kept before the oldest falls off. */
const EDIT_LIMIT = 50;

/** Text/paragraph commands over the selected shape (font/size carry the
 *  combobox value through detail.value; line-spacing's menu items carry
 *  "1.0"-style multipliers). */
const TEXT_FORMAT_COMMANDS: ReadonlySet<string> = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "font-face",
  "font-size",
  "align-left",
  "align-center",
  "align-right",
  "justify",
  "list",
  "numbering",
  "line-spacing",
]);

/** Line-spacing menu multiplier → the lineSpacingPercent the JSON carries. */
const LINE_SPACING_OF: ReadonlyMap<string, number> = new Map([
  ["1.0", 100],
  ["1.5", 150],
  ["2.0", 200],
  ["2.5", 250],
  ["3.0", 300],
]);

/** Ribbon alignment command → TextAlignment value. */
const ALIGNMENTS: ReadonlyMap<string, "left" | "center" | "right" | "justify"> = new Map([
  ["align-left", "left"],
  ["align-center", "center"],
  ["align-right", "right"],
  ["justify", "justify"],
]);

/** Projected paragraph alignment → the CSS text-align the edit overlay uses
 *  (distribute reads as justified — CSS has no separate token). */
const TEXT_ALIGN_OF: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
  both: "justify",
  distribute: "justify",
};

/** The core painter's default ink — colorless runs keep the same gray in
 *  the edit overlay instead of the browser's pure black. */
const TEXT_INK = "#1b1b1b";

/** Browser MIME → the JSON picture type (pptx's supported raster set). */
const PICTURE_TYPES: ReadonlyMap<string, "png" | "jpg" | "gif" | "bmp"> = new Map([
  ["image/png", "png"],
  ["image/jpeg", "jpg"],
  ["image/gif", "gif"],
  ["image/bmp", "bmp"],
]);

const pathsEqual = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

const readFileAsDataURL = (file: File): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

/** One reversible edit: closures over the touched child with the before and
 *  after values captured (no document snapshots — the deck holds media). */
interface DeckEdit {
  undo(): void;
  redo(): void;
}

/** One findable occurrence: the run carrying it, where it sits in the run's
 *  text, and the slide/child the reveal selects. */
interface FindHit {
  slide: number;
  child: number;
  run: TextRunOptions;
  start: number;
}

/** The projected table member's payload as the editor reads it (the layout
 *  type carries `table` as unknown; this is the pptx projection's shape —
 *  see tableMember in @docen/pptx). */
interface TableCellView {
  col: number;
  row: number;
  spanW: number;
  spanH: number;
  anchor?: "top" | "center" | "bottom";
  marginsPx: { left: number; top: number; right: number; bottom: number };
  blocks: LayoutBlock[];
}

interface TableMemberView {
  x: number;
  y: number;
  table: {
    columnWidthsPx: number[];
    rows: { heightPx: number; cells: TableCellView[] }[];
  };
}

/** Commands with a live handler — everything else greys out (the honest
 *  ribbon). Add-in commands join this set at render time. */
const WIRED_COMMANDS: ReadonlySet<string> = new Set([
  "zoom-in",
  "zoom-out",
  "zoom-100",
  "undo",
  "redo",
  "save-as",
  "paste",
  "new-slide",
  "delete-slide",
  "duplicate-slide",
  "text-box",
  "insert-picture",
  "insert-table",
  "shapes",
  "format-background",
  "slide-size",
  "bring-front",
  "send-back",
  "gridlines",
  "find",
  "replace",
  "notes",
  "normal",
  "transition",
  "effect-options",
  "apply-to-all",
  "hyperlink",
  "from-beginning",
  "from-current",
  ...TEXT_FORMAT_COMMANDS,
]);

@customElement({
  name: "docen-presentation",
  template: presentationTemplate,
  styles: presentationStyles,
})
class DocenPresentation extends AddinHost {
  /** The parsed deck — the editable model the gestures write back into. */
  #presJson: PresentationOptions | null = null;
  #pres: ProjectedPresentation | null = null;
  #app: App | null = null;
  #thumbApp: App | null = null;
  #thumbScale = 1;
  #zoom = 100;
  #overlay: DrawingOverlay | null = null;
  /** The selected object: the slide child, or — with `member` set — the
   *  group-nested member it descended into (child indexes under the group). */
  #selection: { slide: number; child: number; member?: number[] } | null = null;
  /** The in-place shape-text editor: the textarea floats over the shape while
   *  it holds the session (its position; the text is read back on exit). */
  #textEditor: HTMLTextAreaElement | null = null;
  /** The table cell under the in-place edit (grid row/col — the session
   *  itself lives in #textEditor; null there means a shape text edit). */
  #tableEdit: { row: number; col: number } | null = null;
  /** Drawing gridlines visibility (a view state — not part of the deck). */
  #gridlines = false;
  #edits: DeckEdit[] = [];
  #editIndex = -1;
  /** The find session: the query the hit list was scanned for and the cursor
   *  in it — the dialog-level highlight PowerPoint's find keeps. */
  #findQuery = "";
  #findCase = false;
  #findHits: FindHit[] = [];
  #findAt = -1;
  /** The slide the notes textarea currently carries (the commit target —
   *  the viewport may move while the session is open). */
  #notesSlideIndex = 0;
  /** The slideshow state: which slide is up and the zoom to restore. */
  #presenting = false;
  #showingSlide = 0;
  #savedZoom = 100;
  #showWheelAt = 0;
  #langObserver?: MutationObserver;
  #unsubLang?: () => void;
  #scrollRaf?: number;

  /** Template refs into the thumbnails panel (see chrome.ts). */
  @observable thumbStrip?: HTMLElement;
  @observable thumbSelection?: HTMLElement;
  @observable gridlines?: HTMLElement;
  /** The speaker-notes pane (chrome.ts): the pane wrapper and its textarea. */
  @observable notesPane?: HTMLElement;
  @observable notesEditor?: HTMLTextAreaElement;

  connectedCallback(): void {
    super.connectedCallback();
    // Forward this host's `lang` attribute to the internal <docen-workspace>
    // so resolveLang honors <docen-presentation lang>, not just <html lang>
    // (same MutationObserver shape as the document element).
    this.#langObserver = new MutationObserver(() => this.#syncLang());
    this.#langObserver.observe(this, { attributes: true, attributeFilter: ["lang"] });
    this.#syncLang();
    void this.#connect();
  }

  async #connect(): Promise<void> {
    await registerComponents();
    if (!this.isConnected) return;
    applyTheme(resolveTheme(this.getAttribute("theme")));
    const root = this.shadowRoot!;
    root.addEventListener("command", this.#onCommand as EventListener);
    root.addEventListener("change", this.#onChange as EventListener);
    root
      .querySelector<HTMLElement>("docen-status-bar")
      ?.addEventListener("zoom:change", this.#onZoomChange as EventListener);
    root
      .querySelector<HTMLInputElement>("#file-input")
      ?.addEventListener("change", this.#onFileChange as EventListener);
    root
      .querySelector<HTMLInputElement>("#picture-input")
      ?.addEventListener("change", this.#onPictureChange as EventListener);
    root
      .querySelector("docen-find-replace-dialog")
      ?.addEventListener("find-replace:action", this.#onFindReplace as EventListener);
    root
      .querySelector("docen-link-dialog")
      ?.addEventListener("link:ok", this.#onLinkOk as EventListener);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#area()?.addEventListener("wheel", this.#onShowWheel, { passive: false });
    this.#area()?.addEventListener("scroll", this.#onScroll);
    this.thumbStrip?.addEventListener("click", this.#onThumbClick);
    this.#stage.addEventListener("pointerdown", this.#onStagePointerDown);
    this.#stage.addEventListener("dblclick", this.#onStageDblClick);
    document.addEventListener("keydown", this.#onKeyDown);
    // The selection frame lives over the slide surface (the canvas element
    // Leafer mounts is recreated per deck render — the overlay survives it).
    this.#overlay = new DrawingOverlay({
      scale: () => this.#zoom / 100,
      applyBox: (box) => {
        // A group member's frame is display-only this round — resize needs
        // the member's scaled geometry write-back, so the drag snaps back.
        if (this.#selection?.member) return;
        this.#applyGesture((child) => resizeChild(child, this.#slideBoxOf(box)));
      },
      applyOffset: (dx, dy) => {
        // A member drag moves in child space: the group scale divides the
        // slide-px delta before it lands on the member's own fields.
        const s = this.#memberScale();
        this.#applyGesture((child) => offsetChild(child, dx / s.sx, dy / s.sy));
      },
      applyRotation: (delta) => {
        if (this.#selection?.member) return;
        this.#applyGesture((child) => rotateChild(child, delta));
      },
    });
    this.#canvasHost().append(this.#overlay.el);
    // Locale switches re-stamp the chrome (header + ribbon labels).
    this.#unsubLang = observeLang(() => this.#renderChrome());
    this.#renderChrome();
    if (this.#pres) this.#renderDeck();
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#langObserver?.disconnect();
    this.#langObserver = undefined;
    this.#unsubLang?.();
    this.#unsubLang = undefined;
    this.shadowRoot?.removeEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot?.removeEventListener("change", this.#onChange as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.removeEventListener("zoom:change", this.#onZoomChange as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLInputElement>("#file-input")
      ?.removeEventListener("change", this.#onFileChange as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLInputElement>("#picture-input")
      ?.removeEventListener("change", this.#onPictureChange as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#onFindReplace as EventListener);
    this.shadowRoot
      ?.querySelector("docen-link-dialog")
      ?.removeEventListener("link:ok", this.#onLinkOk as EventListener);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#area()?.removeEventListener("wheel", this.#onShowWheel);
    this.#area()?.removeEventListener("scroll", this.#onScroll);
    this.thumbStrip?.removeEventListener("click", this.#onThumbClick);
    this.#stage.removeEventListener("pointerdown", this.#onStagePointerDown);
    this.#stage.removeEventListener("dblclick", this.#onStageDblClick);
    document.removeEventListener("keydown", this.#onKeyDown);
    this.#overlay?.el.remove();
    this.#overlay = null;
    if (this.#scrollRaf != null) cancelAnimationFrame(this.#scrollRaf);
    this.#app?.destroy();
    this.#app = null;
    this.#thumbApp?.destroy();
    this.#thumbApp = null;
  }

  /** Parse a .pptx payload, paint every slide, and stamp the chrome (filename
   *  attribute + status bar). Safe before connection — the deck renders on
   *  connect. */
  async openPresentation(data: DataType): Promise<void> {
    const pres = await parsePresentation(data);
    this.#presJson = pres;
    // A fresh deck starts clean: the previous deck's edit closures and
    // selection must not survive the swap.
    this.#exitTextEditing(false);
    this.#selection = null;
    this.#edits = [];
    this.#editIndex = -1;
    // The notes session restarts on the fresh deck's first slide.
    this.#notesSlideIndex = 0;
    this.#syncNotesPane(0);
    this.#pres = projectPresentation(pres);
    // The base class attaches an empty shadowRoot ahead of connect, so
    // shadowRoot presence does not mean the template stamped — same bail
    // #renderChrome does until connectedCallback renders the stored deck.
    if (!this.shadowRoot?.querySelector(".stage")) return;
    this.#renderDeck();
    this.#renderChrome();
  }

  /** Drop the open deck and reset the chrome. */
  closePresentation(): void {
    this.#presJson = null;
    this.#pres = null;
    this.#exitTextEditing(false);
    this.#selection = null;
    this.#edits = [];
    this.#editIndex = -1;
    this.#overlay?.hide();
    this.removeAttribute("filename");
    // The notes session dies with the deck.
    if (this.notesPane) this.notesPane.hidden = true;
    if (this.notesEditor) this.notesEditor.value = "";
    this.#notesSlideIndex = 0;
    if (!this.shadowRoot?.querySelector(".stage")) return;
    this.#app?.destroy();
    this.#app = null;
    this.#thumbApp?.destroy();
    this.#thumbApp = null;
    this.#stage.replaceChildren();
    this.shadowRoot.querySelector<HTMLElement>(".slides-panel")?.setAttribute("hidden", "");
    const bar = this.#statusBar();
    if (bar) {
      bar.removeAttribute("page");
      bar.removeAttribute("total");
      bar.removeAttribute("pageLabel");
    }
    this.#renderChrome();
  }

  // ── Chrome ───────────────────────────────────────────────────────────────

  #area(): HTMLElement | null {
    return this.shadowRoot?.querySelector<HTMLElement>("docen-document-area") ?? null;
  }

  get #stage(): HTMLDivElement {
    return this.shadowRoot!.querySelector<HTMLDivElement>(".stage")!;
  }

  #statusBar(): HTMLElement | null {
    return this.shadowRoot?.querySelector<HTMLElement>("docen-status-bar") ?? null;
  }

  #workspaceEl(): HTMLElement | null {
    return this.shadowRoot?.querySelector("docen-workspace") ?? null;
  }

  /** Stamp the header + ribbon markup for the active locale (re-run on lang
   *  change and on open/close). */
  #renderChrome(): void {
    const root = this.shadowRoot;
    // @attr change callbacks can fire before the template stamps (the
    // shadowRoot exists but is empty) — bail until connectedCallback's
    // explicit call does the first render.
    const titleBar = root?.querySelector("docen-title-bar");
    if (!root || !titleBar) return;
    titleBar.innerHTML = this.#renderHeader();
    // Built-in PowerPoint tabs come from presentationRibbonTabs; external
    // add-ins layer their own tabs on top via mergeRibbonSchema.
    const tabs = [...presentationRibbonTabs(), ...mergeRibbonSchema(this.addins)];
    const ribbonEl = root.querySelector("docen-ribbon")!;
    // The workspace is the i18n scope — labels resolve against
    // <docen-workspace lang> (forwarded from <docen-presentation lang>).
    // closest() can't reach it from the detached fragment, so it's handed in.
    const workspace = root.querySelector("docen-workspace") ?? document.documentElement;
    ribbonEl.replaceChildren(renderRibbonFromSchema(tabs, [], workspace));
    // Feed the ribbon schema to the command search (re-runs per chrome render,
    // the same cadence as the document editor).
    const searchEl = root.querySelector("docen-command-search") as
      | (HTMLElement & { setTabs(tabs: readonly unknown[], scope: Element | null): void })
      | null;
    searchEl?.setTabs(tabs, root.querySelector("docen-workspace"));
    this.notesEditor?.setAttribute("placeholder", t("ppt.notes.placeholder", this));
    this.#applyRibbonGreying();
  }

  #renderHeader(): string {
    const user = this.getAttribute("user") ?? "";
    const avatar = this.getAttribute("avatar") ?? "";
    const filename = this.getAttribute("filename") ?? t("ppt.header.doc-name", this);
    const initial = user.trim().charAt(0).toUpperCase();
    const avatarMarkup = avatar
      ? `<img class="avatar avatar-img" src="${escapeHtml(avatar)}" alt="" />`
      : initial
        ? `<span class="avatar">${initial}</span>`
        : "";
    // QAT undo/redo follow the edit stack; save stays disabled-honest (the
    // file lives on disk — Save As carries the edits out).
    const qat = [
      { id: "save", icon: "save", disabled: true },
      { id: "undo", icon: "undo", disabled: this.#editIndex < 0 },
      { id: "redo", icon: "redo", disabled: this.#editIndex >= this.#edits.length - 1 },
    ]
      .map(
        (c) =>
          `<docen-ribbon-button icon="${c.icon}" label="${t(`ppt.header.${c.id}`, this)}" event="${c.id}" icon-only${c.disabled ? " disabled" : ""}></docen-ribbon-button>`,
      )
      .join("");
    return `
          <div slot="start" style="display:flex;align-items:center;gap:4px">
            <span style="font-weight:600;font-size:13px;padding-inline:6px">${t("ppt.header.brand", this)}</span>
            ${qat}
            <fluent-menu style="margin-inline-start:10px">
              <fluent-menu-button
                slot="trigger"
                appearance="subtle"
                style="max-width:36vw;overflow:hidden;white-space:nowrap"
                title="${escapeHtml(filename)}"
              >${escapeHtml(filename)}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new" disabled>${t("ppt.header.new", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="open">${t("ppt.header.open", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="save-as"${this.#presJson ? "" : " disabled"}>${t("ppt.header.save-as", this)}</fluent-menu-item>
                <fluent-menu-item data-event="print" disabled>${t("ppt.header.print", this)}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="close">${t("ppt.header.close", this)}</fluent-menu-item>
              </fluent-menu-list>
            </fluent-menu>
          </div>
          <docen-command-search slot="search"></docen-command-search>
          <div slot="end" style="display:flex;align-items:center;gap:4px">
            <span style="display:inline-flex;align-items:center;gap:6px;padding-inline:6px">${avatarMarkup}${escapeHtml(user)}</span>
          </div>`;
  }

  /** Grey every ribbon control whose command has no handler (external add-ins
   *  count as handlers via their commands table). */
  #applyRibbonGreying(): void {
    const wired = new Set(WIRED_COMMANDS);
    for (const addin of this.addins) {
      for (const key of Object.keys(addin.commands ?? {})) wired.add(key);
    }
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(
      "docen-ribbon-button[event], docen-ribbon-toggle-button[event], docen-ribbon-split-button[event], docen-ribbon-menu[event]",
    ) ?? []) {
      const event = el.getAttribute("event");
      if (!event || wired.has(event)) continue;
      el.setAttribute("disabled", "");
    }
  }

  // ── Events ───────────────────────────────────────────────────────────────

  /** Title-bar menu items carry their action in `data-event`; the notes
   *  textarea's change (its blur commit) routes to the notes session. */
  readonly #onChange = (event: Event): void => {
    const target = event.target as HTMLElement;
    if (target === this.notesEditor) return this.#commitNotes();
    const name = target?.dataset?.event;
    if (name === "open") this.#pickFile();
    else if (name === "save-as") void this.#saveAs();
    else if (name === "close") this.closePresentation();
  };

  readonly #onCommand = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const name = event.detail?.event;
    if (typeof name !== "string") return;
    // External add-ins take over first; the wired set handles the rest.
    if (this.dispatchCommand(name, event.detail?.value)) return;
    if (name === "zoom-in") this.#setZoom(this.#zoom + 10);
    else if (name === "zoom-out") this.#setZoom(this.#zoom - 10);
    else if (name === "zoom-100") this.#setZoom(100);
    else if (name === "undo") this.#undo();
    else if (name === "redo") this.#redo();
    else if (name === "save-as") void this.#saveAs();
    else if (name === "new-slide") this.#insertSlide();
    else if (name === "delete-slide") this.#deleteSlide();
    else if (name === "duplicate-slide") this.#duplicateSlide();
    else if (name === "text-box") this.#insertTextBox();
    else if (name === "insert-table") this.#insertTable();
    else if (name === "shapes" && event.detail?.value) this.#insertShape(event.detail.value);
    else if (name === "format-background") this.#setBackground(event.detail?.value);
    else if (name === "slide-size") this.#toggleSlideSize();
    else if (name === "insert-picture") this.#pickPicture();
    else if (name === "bring-front") this.#reorderSelected("front");
    else if (name === "send-back") this.#reorderSelected("back");
    else if (name === "gridlines") this.#toggleGridlines();
    else if (name === "find" || name === "replace") this.#findDialog()?.show();
    else if (name === "paste") void this.#pasteFromClipboard();
    else if (name === "notes") this.#toggleNotes();
    else if (name === "normal") this.#enterNormalView();
    else if (name === "transition" && event.detail?.value) this.#setTransition(event.detail.value);
    else if (name === "effect-options" && event.detail?.value) {
      this.#setTransitionSpeed(event.detail.value);
    } else if (name === "apply-to-all") this.#applyTransitionToAll();
    else if (name === "hyperlink") this.#openLinkDialog();
    else if (name === "from-beginning") this.#startShow("beginning");
    else if (name === "from-current") this.#startShow("current");
    else if (TEXT_FORMAT_COMMANDS.has(name)) this.#applyTextFormat(name, event.detail?.value);
  };

  readonly #onZoomChange = (event: Event): void => {
    const zoom = (event as CustomEvent<{ zoom?: number }>).detail?.zoom;
    if (typeof zoom === "number") this.#setZoom(zoom);
  };

  // ── Zoom ─────────────────────────────────────────────────────────────────

  #setZoom(pct: number): void {
    this.#zoom = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.round(pct)));
    this.#applyZoom();
    this.#statusBar()?.setAttribute("zoom", String(this.#zoom));
    this.#syncSlideIndicator();
  }

  /** The zoom IS the layout: the stage's CSS box sizes to the scaled strip
   *  and the tree paints in unzoomed slide px behind a matching scale — no
   *  CSS zoom, so the overlay layer and its handles stay screen-sized and
   *  the canvas bitmap stays 1:1 sharp at every level (the document
   *  editor's zoom model). */
  #applyZoom(): void {
    const s = this.#zoom / 100;
    const pres = this.#pres;
    const strip = pres ? pres.slides.length * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX : 0;
    this.#stage.style.width = `${pres ? pres.widthPx * s : 0}px`;
    this.#stage.style.height = `${strip * s}px`;
    const tree = this.#app?.tree as IGroup | undefined;
    if (tree) tree.scale = { x: s, y: s };
    // PowerPoint's 0.5" grid, in slide px, scaled with the zoom.
    this.gridlines?.style.setProperty(
      "background-size",
      `${GRID_PITCH_PX * s}px ${GRID_PITCH_PX * s}px`,
    );
    if (this.#selection) {
      const hit = this.#selectedBox();
      if (hit) this.#overlay?.refresh(hit.box, hit.rotation);
    }
  }

  // ── Slide strip ──────────────────────────────────────────────────────────

  #onScroll = (): void => {
    if (this.#scrollRaf != null) return;
    this.#scrollRaf = requestAnimationFrame(() => {
      this.#scrollRaf = undefined;
      this.#syncSlideIndicator();
    });
  };

  /** The status bar's slide number follows the viewport center (there is no
   *  caret to track — Word's number follows the caret, PowerPoint's follows
   *  the visible slide). */
  #syncSlideIndicator(): void {
    const pres = this.#pres;
    const area = this.#area();
    const bar = this.#statusBar();
    if (!pres || !area || !bar || pres.slides.length === 0) return;
    const scale = this.#zoom / 100;
    const pitch = (pres.heightPx + SLIDE_GAP_PX) * scale;
    const center = area.scrollTop + area.clientHeight / 2 - SLIDE_GAP_PX * scale;
    const index = Math.max(0, Math.min(pres.slides.length - 1, Math.floor(center / pitch)));
    bar.setAttribute("page", String(index + 1));
    bar.setAttribute("total", String(pres.slides.length));
    bar.setAttribute("pageLabel", t("ppt.status.slide-of", this));
    this.#syncThumbSelection(index);
    this.#syncNotesPane(index);
  }

  /** Repaint the deck. With `slide` set (a single-slide edit), only that
   *  slide's group swaps out — the rest of the strip stays untouched; without
   *  it the whole tree repaints (open, structural slide changes). */
  #renderDeck(slide?: number): void {
    const pres = this.#pres;
    if (!pres) return;
    // One Leafer app per connection; deck repaints clear the tree in place —
    // destroy + recreate blanks the canvas for a frame on every gesture.
    this.#app ??= new App({
      view: this.#stage,
      fill: "transparent",
      tree: { type: "design" },
      move: { disabled: true },
      wheel: { disabled: true },
    });
    const tree = this.#app.tree as unknown as IGroup;
    if (slide !== undefined && tree.children.length === pres.slides.length) {
      repaintSlideAt(
        tree,
        pres,
        slide,
        () => this.#app?.forceRender(),
        pres.heightPx + SLIDE_GAP_PX,
        SLIDE_GAP_PX,
      );
    } else {
      tree.clear();
      paintSlideDeck(tree, pres, () => this.#app?.forceRender());
    }
    this.#applyZoom();
    this.#renderThumbnails(slide);
    this.#syncSlideIndicator();
  }

  // ── Slide-object selection ───────────────────────────────────────────────

  /** The wrapper that anchors the overlay (the stage itself is cleared per
   *  deck render, so the overlay must not live inside it). */
  #canvasHost(): HTMLElement {
    return this.shadowRoot!.querySelector<HTMLElement>(".docen-canvas")!;
  }

  /** Slide-absolute box of the current selection (strip space: slide-local
   *  box plus the slide's offset in the strip) plus its spin. A member
   *  selection shows the member's own box (unspun this round). */
  #selectedBox(): { box: Box; rotation: number } | null {
    const sel = this.#selection;
    const presJson = this.#presJson;
    if (!sel || !presJson) return null;
    const slide = presJson.slides?.[sel.slide];
    const pres = this.#pres!;
    const stripY = sel.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
    if (sel.member) {
      const group = slide?.children?.[sel.child];
      const m = group && memberByPath(group, sel.member);
      if (!m) return null;
      return {
        box: { x: m.x, y: m.y + stripY, width: m.width, height: m.height },
        rotation: 0,
      };
    }
    const hit = slideHits(slide ?? {}).find((h) => h.child === sel.child);
    if (!hit) return null;
    return {
      box: { x: hit.box.x, y: hit.box.y + stripY, width: hit.box.width, height: hit.box.height },
      rotation: hit.rotation,
    };
  }

  /** The slide child at a member path under a group child — undefined when
   *  the path dangles or stops above the leaf. */
  #childAt(child: SlideChild, path: readonly number[] | undefined): SlideChild | undefined {
    let node = child;
    for (const i of path ?? []) {
      if (!("group" in node)) return undefined;
      const next = node.group.children?.[i];
      if (!next) return undefined;
      node = next;
    }
    return node;
  }

  /** The selected slide object: the child itself, or the group-nested member
   *  the selection descended into. */
  #selectedChild(): SlideChild | undefined {
    const sel = this.#selection;
    const child = this.#presJson?.slides?.[sel?.slide ?? -1]?.children?.[sel?.child ?? -1];
    return child ? this.#childAt(child, sel?.member) : undefined;
  }

  /** The child-space → px scale under the current member selection (identity
   *  at top level). */
  #memberScale(): { sx: number; sy: number } {
    const sel = this.#selection;
    if (!sel?.member?.length) return { sx: 1, sy: 1 };
    const group = this.#presJson?.slides?.[sel.slide]?.children?.[sel.child];
    return (group && memberByPath(group, sel.member))?.scale ?? { sx: 1, sy: 1 };
  }

  /** Select a slide object (or clear). The frame shows the strip-space box. */
  #select(sel: { slide: number; child: number; member?: number[] } | null): void {
    if (this.#textEditor) this.#exitTextEditing(true);
    this.#selection = sel;
    if (!sel) return this.#overlay?.hide();
    const hit = this.#selectedBox();
    if (hit) this.#overlay?.show(hit.box, hit.rotation);
    else this.#overlay?.hide();
  }

  /** Double-click on a shape: float a textarea over its text area and hand
   *  the session to it (PowerPoint's in-place text edit). The overlay's
   *  resize handles step aside for the duration. The textarea styles from
   *  the projected member the canvas painted (insets, first-run face, first
   *  paragraph's align, the shape's fill) so the edit state reads like the
   *  render state. */
  #enterTextEditing(): void {
    const sel = this.#selection;
    const pres = this.#pres;
    const presJson = this.#presJson;
    const child = this.#selectedChild();
    const text = child ? shapeTextOf(child) : null;
    if (!sel || !pres || !presJson || !child || text === null) return;
    // The edit anchor: a member's box comes from the group walk, a top-level
    // child's from the slide's hit list.
    let box: Box | null = null;
    if (sel.member) {
      const group = presJson.slides?.[sel.slide]?.children?.[sel.child];
      const m = group && memberByPath(group, sel.member);
      if (m) box = { x: m.x, y: m.y, width: m.width, height: m.height };
    } else {
      const hit = slideHits(presJson.slides?.[sel.slide] ?? {}).find((h) => h.child === sel.child);
      box = hit?.box ?? null;
    }
    if (!box) return;
    const scale = this.#zoom / 100;
    const stripY = box.y + sel.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
    // The projected text-box member for this shape: boxes project with the
    // same slide-absolute geometry (groups fold their affine in), so box
    // equality identifies it.
    const member = pres.slides[sel.slide]?.members.find(
      (m): m is Extract<LayoutDrawingMember, { kind: "textBox" }> => {
        if (m.kind !== "textBox") return false;
        return (
          m.x === box!.x && m.y === box!.y && m.width === box!.width && m.height === box!.height
        );
      },
    );
    // The member's first paragraph (its own blocks are the layout block
    // union; a text box only ever carries paragraphs) and its insets —
    // normalized, the layout type leaves every edge optional.
    const para = member?.blocks[0];
    const first = para?.kind === "paragraph" ? para : undefined;
    // The first text run carries the face the overlay renders (a break-only
    // paragraph falls to the strut style).
    const run =
      first && (first.inline.find((i) => i.kind === "text")?.style ?? first.defaultTextStyle);
    const ins = member?.insets && {
      left: member.insets.left ?? 0,
      top: member.insets.top ?? 0,
      right: member.insets.right ?? 0,
      bottom: member.insets.bottom ?? 0,
    };
    // The painted stack advances lines by the engine's normal line height and
    // hangs shape-text glyphs at 0.85 em below the line top (the painter's
    // shape-text model). The overlay reproduces both: an explicit line-height
    // sized from the same metric source the paint context uses, plus a
    // baseline correction that moves the CSS baseline (half-leading + font
    // ascent) onto the painted depth — CSS offers no direct baseline control.
    const face = run ? familyOfSlot(run.family, false) : "";
    const linePx = run
      ? browserFontMetrics.normalRatio({
          family: face,
          bold: run.bold === true,
          italic: run.italic === true,
        }) * run.sizePx
      : 0;
    const baselineShift = run
      ? this.#baselineShiftOf(
          face,
          run.bold === true,
          run.italic === true,
          run.sizePx,
          linePx,
          scale,
        )
      : 0;
    const editor = document.createElement("textarea");
    editor.className = "shape-text-editor";
    // The rows=2 default inflates scrollHeight to two lines — the slack math
    // (and every grow pass) must measure the real content.
    editor.rows = 1;
    editor.value = text;
    Object.assign(editor.style, {
      left: `${box.x * scale}px`,
      top: `${stripY * scale}px`,
      width: `${box.width * scale}px`,
      height: `${box.height * scale}px`,
      ...(run && ins
        ? {
            padding: `${ins.top * scale}px ${ins.right * scale}px ${ins.bottom * scale}px ${ins.left * scale}px`,
            ...(member?.fill ? { background: `#${member.fill}` } : {}),
            fontFamily: JSON.stringify(run.family),
            fontSize: `${run.sizePx * scale}px`,
            lineHeight: `${linePx * scale}px`,
            textAlign: TEXT_ALIGN_OF[first!.align ?? "left"] ?? "left",
            color: run.color ? `#${run.color}` : TEXT_INK,
            ...(run.bold ? { fontWeight: "bold" } : {}),
            ...(run.italic ? { fontStyle: "italic" } : {}),
          }
        : { fontSize: `${((firstRunSizeOf(child) * 4) / 3) * scale}px` }),
    });
    // Keep the text visible as it grows, and keep the box recognizable as it
    // doesn't: the edit frame starts at the shape's full height (a short text
    // must not shrink the fill/border rectangle under the user — PowerPoint
    // keeps the frame too) and only grows past it when the text overflows.
    // The baseline correction rides the top inset in every branch; top-anchored
    // boxes grow the frame (height only, the width is the shape's), center/
    // bottom re-run the painter's slack math — the stack starts at the top
    // inset plus half (or all) of the leftover inner height, and overflow
    // spills below the box like the painted stack does.
    const frame = 3; // the edit frame's top+bottom borders
    const syncLayout = (): void => {
      if (!member || !ins) {
        editor.style.height = "auto";
        editor.style.height = `${Math.max(box.height * scale, editor.scrollHeight + frame)}px`;
        return;
      }
      editor.style.height = "auto";
      editor.style.paddingTop = `${ins.top * scale + baselineShift}px`;
      if (member.anchor === "top") {
        editor.style.height = `${Math.max(box.height * scale, editor.scrollHeight + frame)}px`;
        return;
      }
      const padTB = (ins.top + ins.bottom) * scale;
      const content = editor.scrollHeight - padTB;
      const slack = member.autoFit ? 0 : box.height * scale - frame - padTB - content;
      if (slack >= 0) {
        editor.style.height = `${box.height * scale}px`;
        editor.style.paddingTop = `${ins.top * scale + baselineShift + (member.anchor === "center" ? slack / 2 : slack)}px`;
      } else {
        editor.style.height = `${content + padTB + frame}px`;
      }
    };
    editor.addEventListener("keydown", (event) => {
      // Escape leaves the edit (committing); typing keys stay in the textarea.
      if (event.key === "Escape") {
        event.stopPropagation();
        this.#exitTextEditing(true);
      }
      event.stopPropagation();
    });
    editor.addEventListener("input", syncLayout);
    this.#canvasHost().append(editor);
    // Measure with the element in the tree — the first pass sizes a
    // top-anchored frame and shifts center/bottom onto the painted baseline.
    syncLayout();
    editor.focus();
    editor.select();
    this.#textEditor = editor;
    this.#overlay?.hide();
  }

  /** Leave the text edit. `write` commits the textarea's text back into the
   *  shape (an undo-able edit when it changed); plain exits just drop it.
   *  A table cell session dispatches to its own exit. */
  #exitTextEditing(write: boolean): void {
    if (this.#tableEdit) return this.#exitTableCellEditing(write);
    const editor = this.#textEditor;
    this.#textEditor = null;
    editor?.remove();
    const sel = this.#selection;
    const child = this.#selectedChild();
    if (editor && write && sel && child && "shape" in child) {
      const shape = child.shape;
      const text = editor.value;
      if (text !== shapeTextOf(child) && shape.textBody) {
        const before = structuredClone(shape.textBody);
        if (writeShapeText(child, text)) {
          const after = structuredClone(shape.textBody);
          this.#pushEdit({
            undo: () => {
              shape.textBody = structuredClone(before);
              this.#reproject(sel.slide);
            },
            redo: () => {
              shape.textBody = structuredClone(after);
              this.#reproject(sel.slide);
            },
          });
          this.#reproject(sel.slide);
        }
      }
    }
    this.#restoreOverlay();
  }

  /** The selection frame returns after the editor steps aside (a committed
   *  write-back refreshes it through #reproject; a no-op exit restores the
   *  pre-edit frame itself). */
  #restoreOverlay(): void {
    if (this.#textEditor || !this.#selection) return;
    const hit = this.#selectedBox();
    if (hit) this.#overlay?.show(hit.box, hit.rotation);
  }

  // ── Table cell editing ─────────────────────────────────────────────────

  /** The selected table's projected member and source table (payload cast —
   *  the layout member carries `table` as unknown). Origin equality
   *  identifies the member: a top-level frame projects from the same x/y the
   *  hit test reads, a group member from the group walk's box. */
  #tableMemberOf(): {
    slide: number;
    table: TableOptions;
    member: TableMemberView;
  } | null {
    const sel = this.#selection;
    const pres = this.#pres;
    const presJson = this.#presJson;
    if (!sel || !pres || !presJson) return null;
    const child = this.#selectedChild();
    if (!child || !("table" in child)) return null;
    let origin: { x: number; y: number } | null = null;
    if (sel.member) {
      const group = presJson.slides?.[sel.slide]?.children?.[sel.child];
      const m = group && memberByPath(group, sel.member);
      if (m) origin = m;
    } else {
      const hit = slideHits(presJson.slides?.[sel.slide] ?? {}).find((h) => h.child === sel.child);
      if (hit) origin = hit.box;
    }
    if (!origin) return null;
    const member = pres.slides[sel.slide]?.members.find(
      (m): m is LayoutDrawingMember & { kind: "table" } =>
        m.kind === "table" && m.x === origin!.x && m.y === origin!.y,
    );
    if (!member) return null;
    return {
      slide: sel.slide,
      table: child.table,
      member: member as unknown as TableMemberView,
    };
  }

  /** The cell whose painted rect contains the slide-local point — bands from
   *  the member's column widths and declared row heights, a span folding its
   *  merged slots onto the origin cell. Returns the cell with its rect. */
  #cellRectAt(
    member: TableMemberView,
    x: number,
    y: number,
  ): { cell: TableCellView; x: number; y: number; width: number; height: number } | null {
    const colX = [0];
    for (const w of member.table.columnWidthsPx) colX.push(colX[colX.length - 1]! + w);
    const rowY = [0];
    for (const row of member.table.rows) rowY.push(rowY[rowY.length - 1]! + row.heightPx);
    const spanEnd = (v: number[], i: number, n: number): number =>
      v[Math.min(i + n, v.length - 1)]!;
    const lx = x - member.x;
    const ly = y - member.y;
    for (const cell of member.table.rows.flatMap((r) => r.cells)) {
      const x0 = colX[Math.min(cell.col, colX.length - 1)]!;
      const y0 = rowY[Math.min(cell.row, rowY.length - 1)]!;
      const x1 = spanEnd(colX, cell.col, cell.spanW);
      const y1 = spanEnd(rowY, cell.row, cell.spanH);
      if (lx >= x0 && lx < x1 && ly >= y0 && ly < y1)
        return { cell, x: member.x + x0, y: member.y + y0, width: x1 - x0, height: y1 - y0 };
    }
    return null;
  }

  /** Float a textarea over one cell of the selected table (PowerPoint's
   *  in-place cell edit): insets, face, paragraph alignment and vertical
   *  anchor come from the painted cell; the text writes back on exit, and a
   *  click on another cell moves the session there. */
  #enterTableCellEditing(x: number, y: number): void {
    const found = this.#tableMemberOf();
    if (!found) return;
    const rect = this.#cellRectAt(found.member, x, y);
    if (!rect) return;
    const cell = rect.cell;
    const source = tableGridOf(found.table).origins.find(
      (o) => o.row === cell.row && o.col === cell.col,
    )?.cell;
    if (!source) return;
    const pres = this.#pres!;
    const scale = this.#zoom / 100;
    const stripY = rect.y + found.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
    // The cell's first text run styles the editor — the same projected-block
    // read the shape editor does.
    const para = cell.blocks[0];
    const first = para?.kind === "paragraph" ? para : undefined;
    const run =
      first && (first.inline.find((i) => i.kind === "text")?.style ?? first.defaultTextStyle);
    const face = run ? familyOfSlot(run.family, false) : "";
    const linePx = run
      ? browserFontMetrics.normalRatio({
          family: face,
          bold: run.bold === true,
          italic: run.italic === true,
        }) * run.sizePx
      : 0;
    const baselineShift = run
      ? this.#baselineShiftOf(
          face,
          run.bold === true,
          run.italic === true,
          run.sizePx,
          linePx,
          scale,
        )
      : 0;
    const m = cell.marginsPx;
    const anchor = cell.anchor ?? "top";
    const editor = document.createElement("textarea");
    editor.className = "shape-text-editor";
    editor.rows = 1;
    editor.value = cellTextOf(source);
    Object.assign(editor.style, {
      left: `${rect.x * scale}px`,
      top: `${stripY * scale}px`,
      width: `${rect.width * scale}px`,
      height: `${rect.height * scale}px`,
      padding: `${m.top * scale}px ${m.right * scale}px ${m.bottom * scale}px ${m.left * scale}px`,
      ...(run
        ? {
            fontFamily: JSON.stringify(face),
            fontSize: `${run.sizePx * scale}px`,
            lineHeight: `${linePx * scale}px`,
            textAlign: TEXT_ALIGN_OF[first!.align ?? "left"] ?? "left",
            color: run.color ? `#${run.color}` : TEXT_INK,
            ...(run.bold ? { fontWeight: "bold" } : {}),
            ...(run.italic ? { fontStyle: "italic" } : {}),
          }
        : { fontSize: `${((firstCellRunSizeOf(source) * 4) / 3) * scale}px` }),
    });
    // Overflow grows the frame downward (the top-anchored shape rule); a
    // center/bottom cell re-runs the painter's slack math so the stack sits
    // where the paint puts it. On commit the row grows to fit — the same
    // growth every re-projection applies.
    const frame = 3; // the edit frame's top+bottom borders
    const syncLayout = (): void => {
      editor.style.height = "auto";
      editor.style.paddingTop = `${m.top * scale + baselineShift}px`;
      const boxH = rect.height * scale;
      if (anchor === "top") {
        editor.style.height = `${Math.max(boxH, editor.scrollHeight + frame)}px`;
        return;
      }
      const padTB = (m.top + m.bottom) * scale;
      const content = editor.scrollHeight - padTB;
      const slack = boxH - frame - padTB - content;
      if (slack >= 0) {
        editor.style.height = `${boxH}px`;
        editor.style.paddingTop = `${m.top * scale + baselineShift + (anchor === "center" ? slack / 2 : slack)}px`;
      } else {
        editor.style.height = `${content + padTB + frame}px`;
      }
    };
    editor.addEventListener("keydown", (event) => {
      // Escape leaves the edit (committing); typing keys stay in the textarea.
      if (event.key === "Escape") {
        event.stopPropagation();
        this.#exitTextEditing(true);
      }
      event.stopPropagation();
    });
    editor.addEventListener("input", syncLayout);
    this.#canvasHost().append(editor);
    syncLayout();
    editor.focus();
    editor.select();
    this.#textEditor = editor;
    this.#tableEdit = { row: cell.row, col: cell.col };
    this.#overlay?.hide();
  }

  /** The source cell a table edit session points at — the grid walk maps the
   *  session's row/col slot back to the cell object. */
  #editedCell(at = this.#tableEdit): TableCellOptions | null {
    if (!at) return null;
    const child = this.#selectedChild();
    if (!child || !("table" in child)) return null;
    return (
      tableGridOf(child.table).origins.find((o) => o.row === at.row && o.col === at.col)?.cell ??
      null
    );
  }

  /** Leave the cell edit: `write` commits the textarea's text into the cell
   *  (an undo-able edit when it changed — the cell mutates in place, so the
   *  undo pair swaps its own keys back). */
  #exitTableCellEditing(write: boolean): void {
    const editor = this.#textEditor;
    const at = this.#tableEdit;
    this.#textEditor = null;
    this.#tableEdit = null;
    editor?.remove();
    const sel = this.#selection;
    const cell = editor && write ? this.#editedCell(at) : null;
    if (cell) {
      const text = editor!.value;
      if (text !== cellTextOf(cell)) {
        const before = structuredClone(cell);
        writeCellText(cell, text);
        const after = structuredClone(cell);
        const restore = (snap: TableCellOptions): void => {
          for (const key of Object.keys(cell) as (keyof TableCellOptions)[]) delete cell[key];
          Object.assign(cell, snap);
        };
        this.#pushEdit({
          undo: () => {
            restore(before);
            this.#reproject(sel!.slide);
          },
          redo: () => {
            restore(after);
            this.#reproject(sel!.slide);
          },
        });
        this.#reproject(sel!.slide);
      }
    }
    this.#restoreOverlay();
  }

  /** A click on another cell of the table under edit: commit the current
   *  cell and float the editor over the new one (PowerPoint's cell hop). */
  #moveTableCellEditing(x: number, y: number): void {
    const found = this.#tableMemberOf();
    const rect = found ? this.#cellRectAt(found.member, x, y) : null;
    if (!rect) return this.#exitTextEditing(true);
    if (this.#tableEdit?.row === rect.cell.row && this.#tableEdit.col === rect.cell.col) return;
    this.#exitTextEditing(true);
    this.#enterTableCellEditing(x, y);
  }

  /** The padding-top correction that moves a textarea's CSS baseline onto
   *  the painted depth: the painted stack hangs shape-text glyphs
   *  leaferBaselinePadPx below the line top while CSS centers the
   *  half-leading box. 0 when the browser can't measure. */
  #baselineShiftOf(
    face: string,
    bold: boolean,
    italic: boolean,
    sizePx: number,
    linePx: number,
    scale: number,
  ): number {
    const probe = document.createElement("canvas").getContext("2d");
    if (!probe) return 0;
    probe.font = `${italic ? "italic " : ""}${bold ? "700 " : ""}${sizePx * scale}px ${JSON.stringify(face)}`;
    const m = probe.measureText("Ag");
    if (!(m.fontBoundingBoxAscent > 0)) return 0;
    const cssBaseline =
      (linePx * scale - m.fontBoundingBoxAscent - m.fontBoundingBoxDescent) / 2 +
      m.fontBoundingBoxAscent;
    return leaferBaselinePadPx(sizePx) * scale - cssBaseline;
  }

  /** Re-project and repaint after a source edit (gestures, delete, undo,
   *  redo — the one paint path every mutation funnels through). `slide`
   *  scopes the repaint to the one slide the edit touched; structural edits
   *  (slide insert/delete) omit it and repaint the whole deck. */
  #reproject(slide?: number): void {
    if (!this.#presJson) return;
    this.#pres = projectPresentation(this.#presJson);
    this.#renderDeck(slide);
    // The paint restarted under the same selection — the frame snaps to the
    // current geometry (and rejects itself if the box vanished).
    const hit = this.#selectedBox();
    if (hit) this.#overlay?.refresh(hit.box, hit.rotation);
    else this.#select(null);
  }

  /** Strip-space box → slide-local: drop the slide's strip offset (the
   *  inverse of #selectedBox) so resize write-back lands on the slide. */
  #slideBoxOf(box: Box): Box {
    const pitch = (this.#pres?.heightPx ?? 0) + SLIDE_GAP_PX;
    const offset = (this.#selection?.slide ?? 0) * pitch + SLIDE_GAP_PX;
    return { ...box, y: box.y - offset };
  }

  /** Record a reversible edit; a new edit truncates the redo branch. */
  #pushEdit(edit: DeckEdit): void {
    this.#edits.length = this.#editIndex + 1;
    this.#edits.push(edit);
    if (this.#edits.length > EDIT_LIMIT) this.#edits.shift();
    this.#editIndex = this.#edits.length - 1;
    this.#syncQat();
  }

  /** Commit a drag gesture to the selected object (a child or a group
   *  member), record it, and repaint. */
  #applyGesture(edit: (child: SlideChild) => void): void {
    const sel = this.#selection;
    const child = this.#selectedChild();
    if (!sel || !child) return;
    const before = captureGeometry(child);
    edit(child);
    const after = captureGeometry(child);
    this.#pushEdit({
      undo: () => {
        restoreGeometry(child, before);
        this.#reproject(sel.slide);
      },
      redo: () => {
        restoreGeometry(child, after);
        this.#reproject(sel.slide);
      },
    });
    this.#reproject(sel.slide);
    // Re-frame the overlay on the committed geometry: a member drag displays
    // the raw slide delta, but the write-back divides it by the group scale,
    // so the drag-end frame no longer matches where the member landed.
    this.#select(sel);
  }

  /** Remove the selected object; undo re-splices the same child back. */
  #deleteSelected(): void {
    const sel = this.#selection;
    const slide = this.#presJson?.slides?.[sel?.slide ?? -1];
    const children = slide?.children;
    if (!sel || !children) return;
    const [child] = children.splice(sel.child, 1);
    if (!child) return;
    this.#select(null);
    this.#pushEdit({
      undo: () => {
        children.splice(sel.child, 0, child);
        this.#reproject(sel.slide);
      },
      redo: () => {
        children.splice(sel.child, 1);
        this.#reproject(sel.slide);
      },
    });
    this.#reproject(sel.slide);
  }

  #undo(): void {
    if (this.#editIndex < 0) return;
    // An in-flight text edit commits first, so undo revokes it (the freshest
    // step) rather than the step before — the user's intent either way.
    this.#exitTextEditing(true);
    this.#edits[this.#editIndex--]!.undo();
    this.#syncQat();
  }

  #redo(): void {
    if (this.#editIndex >= this.#edits.length - 1) return;
    this.#exitTextEditing(true);
    this.#edits[++this.#editIndex]!.redo();
    this.#syncQat();
  }

  // ── Ribbon commands ──────────────────────────────────────────────────────

  /** The slide ribbon commands target: the selected object's slide, else the
   *  slide at the viewport center (the status bar's rule). */
  #activeSlideIndex(): number {
    const pres = this.#pres;
    if (!pres || pres.slides.length === 0) return 0;
    if (this.#selection) return this.#selection.slide;
    const area = this.#area();
    if (!area) return 0;
    const scale = this.#zoom / 100;
    const pitch = (pres.heightPx + SLIDE_GAP_PX) * scale;
    const center = area.scrollTop + area.clientHeight / 2 - SLIDE_GAP_PX * scale;
    return Math.max(0, Math.min(pres.slides.length - 1, Math.floor(center / pitch)));
  }

  /** New blank slide after the active one. Structural — the selection drops
   *  (child indices stay per-slide, but the viewport lands on the new page). */
  #insertSlide(): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const at = this.#activeSlideIndex();
    insertSlideAt(presJson, at);
    this.#select(null);
    this.#pushEdit({
      undo: () => {
        presJson.slides?.splice(at + 1, 1);
        this.#select(null);
        this.#reproject();
      },
      redo: () => {
        insertSlideAt(presJson, at);
        this.#select(null);
        this.#reproject();
      },
    });
    this.#reproject();
    this.#revealSlide(at + 1);
  }

  /** Remove the active slide (the last one stays — the painter has no empty
   *  deck to draw). Structural: full repaint, selection drops, and undo
   *  re-splices the same slide object back at its index. */
  #deleteSlide(): void {
    const presJson = this.#presJson;
    if (!presJson || (presJson.slides?.length ?? 0) <= 1) return;
    const at = this.#activeSlideIndex();
    const removed = deleteSlideAt(presJson, at);
    if (!removed) return;
    this.#select(null);
    this.#pushEdit({
      undo: () => {
        presJson.slides?.splice(at, 0, removed);
        this.#select(null);
        this.#reproject();
        this.#revealSlide(at);
      },
      redo: () => {
        presJson.slides?.splice(at, 1);
        this.#select(null);
        this.#reproject();
        this.#revealSlide(Math.max(0, at - 1));
      },
    });
    this.#reproject();
    this.#revealSlide(Math.max(0, at - 1));
  }

  /** Deep-clone the active slide right after it; undo removes the copy. */
  #duplicateSlide(): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const at = this.#activeSlideIndex();
    const copyAt = duplicateSlideAt(presJson, at);
    if (copyAt < 0) return;
    const copy = presJson.slides?.[copyAt];
    this.#select(null);
    this.#pushEdit({
      undo: () => {
        presJson.slides?.splice(copyAt, 1);
        this.#select(null);
        this.#reproject();
        this.#revealSlide(at);
      },
      redo: () => {
        if (copy) presJson.slides?.splice(copyAt, 0, copy);
        this.#select(null);
        this.#reproject();
        this.#revealSlide(copyAt);
      },
    });
    this.#reproject();
    this.#revealSlide(copyAt);
  }

  /** Move the selected object to the top (front) or bottom (back) of its
   *  slide's z order; the selection follows the object. */
  #reorderSelected(to: "front" | "back"): void {
    const sel = this.#selection;
    const children = this.#presJson?.slides?.[sel?.slide ?? -1]?.children;
    if (!sel || !children || children.length < 2) return;
    const from = sel.child;
    const moved = to === "front" ? children.length - 1 : 0;
    reorderChild(children, from, moved);
    this.#select({ slide: sel.slide, child: moved });
    this.#pushEdit({
      undo: () => {
        reorderChild(children, moved, from);
        this.#select({ slide: sel.slide, child: from });
        this.#reproject(sel.slide);
      },
      redo: () => {
        reorderChild(children, from, moved);
        this.#select({ slide: sel.slide, child: moved });
        this.#reproject(sel.slide);
      },
    });
    this.#reproject(sel.slide);
  }

  /** Append a child on top of the active slide and select it; undo splices
   *  it back out by index (per-slide, so slide-level inserts can't shift it). */
  #insertChild(child: SlideChild): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const slide = this.#activeSlideIndex();
    const host = presJson.slides?.[slide];
    if (!host) return;
    const children = (host.children ??= []);
    const index = children.length;
    children.push(child);
    this.#pushEdit({
      undo: () => {
        children.splice(index, 1);
        this.#select(null);
        this.#reproject(slide);
      },
      redo: () => {
        children.splice(index, 0, child);
        this.#select({ slide, child: index });
        this.#reproject(slide);
      },
    });
    this.#select({ slide, child: index });
    this.#reproject(slide);
  }

  #insertTextBox(): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeTextBox(pres.widthPx, pres.heightPx));
  }

  #insertTable(): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeTable(pres.widthPx, pres.heightPx));
  }

  /** The Shapes gallery's pick: the value is the prstGeom token. */
  #insertShape(geometry: string): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeShape(pres.widthPx, pres.heightPx, geometry as ShapeType));
  }

  /** The color picker's pick: paint the active slide's background solid.
   *  Swatches arrive as color:RRGGBB; anything else (no-fill tokens) has no
   *  background meaning yet and is dropped. */
  #setBackground(value?: string): void {
    const presJson = this.#presJson;
    if (!presJson || !value?.startsWith("color:")) return;
    const color = value.slice(6).toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(color)) return;
    const slide = this.#activeSlideIndex();
    const host = presJson.slides?.[slide];
    if (!host) return;
    const before = host.background;
    const after = { fill: { type: "solid", color } } as SlideOptions["background"];
    host.background = structuredClone(after);
    this.#pushEdit({
      undo: () => {
        host.background = before;
        this.#reproject(slide);
      },
      redo: () => {
        host.background = structuredClone(after);
        this.#reproject(slide);
      },
    });
    this.#reproject(slide);
  }

  /** Cycle the deck between the two named sizes (16:9 ↔ 4:3). Content keeps
   *  its absolute geometry — the canvas re-frames around it. */
  #toggleSlideSize(): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const before = presJson.size ?? "16:9";
    const after = before === "16:9" ? "4:3" : "16:9";
    presJson.size = after;
    this.#pushEdit({
      undo: () => {
        presJson.size = before;
        this.#reproject();
        this.#applyZoom();
      },
      redo: () => {
        presJson.size = after;
        this.#reproject();
        this.#applyZoom();
      },
    });
    this.#reproject();
    this.#applyZoom();
  }

  // ── Find and replace ─────────────────────────────────────────────────────

  #findDialog(): (Element & { show(): void }) | null {
    return this.shadowRoot?.querySelector("docen-find-replace-dialog") ?? null;
  }

  /** Every occurrence of the query in the deck's live text — shapes' text
   *  bodies and table cells (the two writable text homes). Matching stays
   *  within one run; a stretch split across runs is a later batch. */
  #scanFind(query: string, caseSensitive: boolean): FindHit[] {
    const needle = caseSensitive ? query : query.toLowerCase();
    const hits: FindHit[] = [];
    if (!needle) return hits;
    this.#presJson?.slides?.forEach((slide, slideIndex) => {
      (slide.children ?? []).forEach((child, childIndex) => {
        for (const run of childRunsOf(child)) {
          const text = run.text ?? "";
          const hay = caseSensitive ? text : text.toLowerCase();
          for (
            let at = hay.indexOf(needle);
            at >= 0;
            at = hay.indexOf(needle, at + needle.length)
          ) {
            hits.push({ slide: slideIndex, child: childIndex, run, start: at });
          }
        }
      });
    });
    return hits;
  }

  readonly #onFindReplace = (event: Event): void => {
    const detail = (event as CustomEvent).detail as {
      action?: string;
      find?: string;
      replace?: string;
      caseSensitive?: boolean;
    };
    const query = detail.find ?? "";
    const replacement = detail.replace ?? "";
    const caseSensitive = detail.caseSensitive === true;
    if (detail.action === "find-next") this.#findNext(query, caseSensitive);
    else if (detail.action === "replace-next") this.#replaceNext(query, replacement, caseSensitive);
    else if (detail.action === "replace-all") this.#replaceAll(query, replacement, caseSensitive);
  };

  /** Advance to the next occurrence (wrapping), rescanning when the query or
   *  the case option moved off the cached scan; select the carrying object
   *  and bring its slide into view. */
  #findNext(query: string, caseSensitive: boolean): void {
    if (query !== this.#findQuery || caseSensitive !== this.#findCase) {
      this.#findQuery = query;
      this.#findCase = caseSensitive;
      this.#findHits = this.#scanFind(query, caseSensitive);
      this.#findAt = -1;
    }
    if (this.#findHits.length === 0) return;
    this.#findAt = (this.#findAt + 1) % this.#findHits.length;
    this.#revealHit(this.#findHits[this.#findAt]!);
  }

  /** Replace the highlighted occurrence and land the cursor on the next one;
   *  with no live highlight — or one an edit has moved — the first match is
   *  the one replaced. */
  #replaceNext(query: string, replacement: string, caseSensitive: boolean): void {
    this.#findQuery = query;
    this.#findCase = caseSensitive;
    this.#findHits = this.#scanFind(query, caseSensitive);
    if (this.#findHits.length === 0) {
      this.#findAt = -1;
      return;
    }
    // A cached cursor only counts when its stretch still matches (the deck
    // may have been edited since the scan).
    const needle = caseSensitive ? query : query.toLowerCase();
    const live = this.#findHits[this.#findAt];
    const text = live?.run.text ?? "";
    const hay = caseSensitive ? text : text.toLowerCase();
    if (!live || hay.slice(live.start, live.start + needle.length) !== needle) this.#findAt = 0;
    const hit = this.#findHits[Math.min(this.#findAt, this.#findHits.length - 1)]!;
    const anchor = { slide: hit.slide, child: hit.child };
    this.#replaceHits([hit], query, replacement);
    this.#findHits = this.#scanFind(query, caseSensitive);
    // The cursor continues after the replacement — a query contained in its
    // own replacement must not re-catch the stretch just written.
    this.#seekFind(anchor, { run: hit.run, until: hit.start + replacement.length });
  }

  /** Replace every occurrence in one undo step. */
  #replaceAll(query: string, replacement: string, caseSensitive: boolean): void {
    const hits = this.#scanFind(query, caseSensitive);
    if (hits.length === 0) return;
    this.#findQuery = query;
    this.#findCase = caseSensitive;
    this.#replaceHits(hits, query, replacement);
    this.#findHits = [];
    this.#findAt = -1;
  }

  /** Write the replacement at each hit (back-to-front, so later starts in the
   *  same run stay valid) as one reversible edit and repaint the touched
   *  slides. Only the touched runs are snapshotted — the deck holds media. */
  #replaceHits(hits: FindHit[], query: string, replacement: string): void {
    const touched = new Map<TextRunOptions, { before: string; after: string }>();
    for (const hit of [...hits].reverse()) {
      const text = hit.run.text ?? "";
      hit.run.text = text.slice(0, hit.start) + replacement + text.slice(hit.start + query.length);
      // The first touch of a run records its original text; later touches in
      // the same run only refresh the result.
      touched.set(hit.run, { before: touched.get(hit.run)?.before ?? text, after: hit.run.text });
    }
    const slides = new Set(hits.map((hit) => hit.slide));
    const edits = [...touched];
    const apply = (side: "before" | "after"): void => {
      for (const [run, touch] of edits) run.text = touch[side];
      for (const slide of slides) this.#reproject(slide);
    };
    this.#pushEdit({ undo: () => apply("before"), redo: () => apply("after") });
    apply("after");
  }

  /** Put the cursor on the first hit at or after the anchor child (wrapping
   *  to the list's head) and reveal it. `except` drops one run's stretch —
   *  the replacement a replace-next just wrote. */
  #seekFind(
    anchor: { slide: number; child: number },
    except?: { run: TextRunOptions; until: number },
  ): void {
    this.#findAt = this.#findHits.findIndex((hit) => {
      if (hit.slide < anchor.slide) return false;
      if (hit.slide === anchor.slide && hit.child < anchor.child) return false;
      if (except && hit.run === except.run && hit.start < except.until) return false;
      return true;
    });
    if (this.#findAt < 0) this.#findAt = this.#findHits.length - 1;
    const hit = this.#findHits[this.#findAt];
    if (hit) this.#revealHit(hit);
  }

  #revealHit(hit: FindHit): void {
    this.#select({ slide: hit.slide, child: hit.child });
    this.#revealSlide(hit.slide);
  }

  /** Toggle the drawing gridlines — a pure view state, no undo step. */
  #toggleGridlines(): void {
    this.#gridlines = !this.#gridlines;
    if (this.gridlines) {
      this.gridlines.hidden = !this.#gridlines;
      if (this.#gridlines) {
        const s = this.#zoom / 100;
        this.gridlines.style.setProperty(
          "background-size",
          `${GRID_PITCH_PX * s}px ${GRID_PITCH_PX * s}px`,
        );
      }
    }
  }

  // ── Notes pane ───────────────────────────────────────────────────────────

  /** Show/hide the speaker-notes pane (a view state — the deck is untouched
   *  until the textarea commits). */
  #toggleNotes(): void {
    const pane = this.notesPane;
    if (!pane) return;
    pane.hidden = !pane.hidden;
    if (!pane.hidden) this.#syncNotesPane(this.#activeSlideIndex());
  }

  /** The Normal view: back to the default surface — the notes pane closes
   *  (later view states fold in here). */
  #enterNormalView(): void {
    if (this.notesPane) this.notesPane.hidden = true;
  }

  /** Point the textarea at the slide: a no-op while the session is open (the
   *  mid-edit viewport moves must not clobber the text under the caret). */
  #syncNotesPane(index: number): void {
    const pane = this.notesPane;
    const editor = this.notesEditor;
    if (!pane || !editor || pane.hidden) return;
    if (document.activeElement === editor) return;
    this.#notesSlideIndex = index;
    editor.value = slideNotesOf(this.#presJson?.slides?.[index] ?? {});
  }

  /** The textarea's blur commit: write the plain text back to the slide it
   *  was opened for (no repaint — notes don't project onto the canvas) as
   *  one reversible edit. */
  #commitNotes(): void {
    const slide = this.#presJson?.slides?.[this.#notesSlideIndex];
    const editor = this.notesEditor;
    if (!slide || !editor) return;
    const before = slideNotesOf(slide);
    const after = editor.value;
    if (after === before) return;
    writeSlideNotes(slide, after);
    this.#pushEdit({
      undo: () => {
        writeSlideNotes(slide, before);
        this.#syncNotesPane(this.#notesSlideIndex);
      },
      redo: () => {
        writeSlideNotes(slide, after);
        this.#syncNotesPane(this.#notesSlideIndex);
      },
    });
    this.#syncNotesPane(this.#notesSlideIndex);
  }

  // ── Transitions ──────────────────────────────────────────────────────────

  /** The gallery's pick on the active slide: the effect token lands as the
   *  slide's transition; none clears it (playback-only — nothing paints). */
  #setTransition(value: string): void {
    const host = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!host) return;
    const before = host.transition;
    const after: TransitionOptions = { type: value as TransitionType };
    // Picking the live effect again records nothing.
    if (typeof before === "object" ? before?.type === value : value === "none") return;
    if (value === "none") delete host.transition;
    else host.transition = structuredClone(after);
    this.#pushEdit({
      undo: () => {
        if (before === undefined) delete host.transition;
        else host.transition = structuredClone(before);
      },
      redo: () => {
        host.transition = structuredClone(after);
      },
    });
  }

  /** Effect options' speed pick on the active slide's transition. A slide
   *  without a live transition (or carrying a verbatim extension string)
   *  has nothing to tune. */
  #setTransitionSpeed(value: string): void {
    const host = this.#presJson?.slides?.[this.#activeSlideIndex()];
    const speed = value === "slow" || value === "fast" ? value : "medium";
    if (!host || typeof host.transition !== "object") return;
    const before = host.transition;
    const after: TransitionOptions = { ...before, speed };
    if (JSON.stringify(after) === JSON.stringify(before)) return;
    host.transition = structuredClone(after);
    this.#pushEdit({
      undo: () => {
        host.transition = structuredClone(before);
      },
      redo: () => {
        host.transition = structuredClone(after);
      },
    });
  }

  /** Copy the active slide's transition onto every slide (PowerPoint's
   *  Apply To All) as one reversible edit. */
  #applyTransitionToAll(): void {
    const presJson = this.#presJson;
    const slides = presJson?.slides;
    if (!slides) return;
    const after = structuredClone(slides[this.#activeSlideIndex()]?.transition);
    const before = slides.map((slide) => structuredClone(slide.transition));
    const restore = (values: (TransitionOptions | string | undefined)[]): void => {
      slides.forEach((slide, i) => {
        const value = values[i];
        if (value === undefined) delete slide.transition;
        else slide.transition = value;
      });
    };
    this.#pushEdit({
      undo: () => restore(before),
      redo: () => restore(slides.map(() => structuredClone(after))),
    });
    restore(slides.map(() => after));
  }

  // ── Slide show ───────────────────────────────────────────────────────────

  /** Enter the presenting state: the chrome steps aside, the zoom fits one
   *  slide to the viewport, and the browser goes fullscreen when allowed (a
   *  denied request only means the window stays windowed). */
  #startShow(from: "beginning" | "current"): void {
    const pres = this.#pres;
    if (!pres || this.#presenting) return;
    this.#exitTextEditing(true);
    this.#presenting = true;
    this.#savedZoom = this.#zoom;
    this.#workspaceEl()?.classList.add("presenting");
    const area = this.#area();
    if (area) {
      const fit = Math.floor(((area.clientHeight - SLIDE_GAP_PX) / pres.heightPx) * 100);
      this.#setZoom(Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, fit)));
    }
    this.#gotoShowSlide(from === "beginning" ? 0 : this.#activeSlideIndex());
    void this.#workspaceEl()
      ?.requestFullscreen?.()
      .catch(() => {});
  }

  /** Leave the presenting state and restore the zoom and viewport. */
  #endShow(): void {
    if (!this.#presenting) return;
    this.#presenting = false;
    this.#workspaceEl()?.classList.remove("presenting");
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => {});
    this.#setZoom(this.#savedZoom);
    this.#revealSlide(this.#showingSlide);
  }

  /** Jump to a slide (clamped) with an instant scroll — the slideshow's own
   *  paging has no smooth travel. */
  #gotoShowSlide(index: number): void {
    const pres = this.#pres;
    const area = this.#area();
    if (!pres || !area) return;
    this.#showingSlide = Math.max(0, Math.min(pres.slides.length - 1, index));
    const pitch = (pres.heightPx + SLIDE_GAP_PX) * (this.#zoom / 100);
    area.scrollTo({ top: this.#showingSlide * pitch, behavior: "instant" as ScrollBehavior });
    this.#syncSlideIndicator();
  }

  /** Wheel paging in the show, throttled — one notch is one slide. */
  #onShowWheel = (event: WheelEvent): void => {
    if (!this.#presenting) return;
    event.preventDefault();
    const now = performance.now();
    if (now - this.#showWheelAt < 300) return;
    this.#showWheelAt = now;
    this.#gotoShowSlide(this.#showingSlide + (event.deltaY > 0 ? 1 : -1));
  };

  /** The browser's Escape leaves fullscreen ahead of us — fold the show up
   *  with it instead of stranding a windowed show behind a stale state. */
  readonly #onFullscreenChange = (): void => {
    if (!document.fullscreenElement) this.#endShow();
  };

  #pickPicture(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#picture-input")?.click();
  }

  /** File-input picture path: read the file as a data URL, measure its
   *  natural size, and drop a centered picture child on the active slide. */
  readonly #onPictureChange = async (event: Event): Promise<void> => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (file) await this.#insertImageFile(file);
  };

  /** Read an image payload as a data URL, measure it, and drop a centered
   *  picture child on the active slide — the file-input and clipboard paths
   *  converge here. */
  async #insertImageFile(file: File): Promise<void> {
    const pres = this.#pres;
    const type = PICTURE_TYPES.get(file.type);
    if (!pres || !type) return;
    const data = await readFileAsDataURL(file);
    const image = new Image();
    image.src = data;
    try {
      await image.decode();
    } catch {
      return;
    }
    this.#insertChild(
      makePicture(pres.widthPx, pres.heightPx, image.naturalWidth, image.naturalHeight, data, type),
    );
  }

  /** The paste command: the clipboard's leading image lands as a picture
   *  child, plain text as a seeded text box. A denied read is silent — the
   *  browser's permission state, not ours, owns the failure. */
  async #pasteFromClipboard(): Promise<void> {
    const pres = this.#pres;
    if (!pres) return;
    const readClipboard = async (): Promise<{ image?: File; text?: string } | null> => {
      try {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          const mime = item.types.find((t) => PICTURE_TYPES.has(t));
          if (mime)
            return { image: new File([await item.getType(mime)], "clipboard", { type: mime }) };
          if (item.types.includes("text/plain")) {
            const text = (await (await item.getType("text/plain")).text()).slice(0, 10000);
            return { text };
          }
        }
        return null;
      } catch {
        try {
          return { text: (await navigator.clipboard.readText()).slice(0, 10000) || undefined };
        } catch {
          return null;
        }
      }
    };
    const payload = await readClipboard();
    if (!payload) return;
    if (payload.image) await this.#insertImageFile(payload.image);
    else if (payload.text)
      this.#insertChild(makeTextBox(pres.widthPx, pres.heightPx, payload.text));
  }

  /** Apply a font/paragraph command to the text under edit: the cell a table
   *  edit session was in, else the selected shape's text body. The undo pair
   *  swaps the whole target back (geometry snapshots don't reach text); a
   *  command that changed nothing records nothing. */
  #applyTextFormat(name: string, value?: string): void {
    // Format lands on the committed text: an in-flight edit writes back
    // first, or the exit-time write-back would clobber the format.
    const at = this.#tableEdit;
    this.#exitTextEditing(true);
    const sel = this.#selection;
    const child = this.#selectedChild();
    const cell = at ? this.#editedCell(at) : null;
    const body = cell || !child || !("shape" in child) ? null : (child.shape.textBody ?? null);
    if (!cell && !body) return;
    const target = cell ?? body!;
    const paragraphs = cell ? cellParagraphsOf(cell) : bodyParagraphsOf(body!);
    const before = structuredClone(target);
    switch (name) {
      case "bold":
        toggleRunFlag(paragraphs, "bold");
        break;
      case "italic":
        toggleRunFlag(paragraphs, "italic");
        break;
      case "underline":
        toggleRunStyle(paragraphs, "underline", "single");
        break;
      case "strike":
        toggleRunStyle(paragraphs, "strike", "singleStrike");
        break;
      case "font-face":
        if (value) setRunFont(paragraphs, value);
        break;
      case "font-size": {
        const size = Number(value);
        if (Number.isFinite(size) && size > 0) setRunSize(paragraphs, size);
        break;
      }
      default: {
        const alignment = ALIGNMENTS.get(name);
        if (alignment) setParagraphAlignment(paragraphs, alignment);
        else if (name === "list") toggleBullet(paragraphs);
        else if (name === "numbering") toggleNumbering(paragraphs);
        else if (name === "line-spacing") {
          const percent = LINE_SPACING_OF.get(value ?? "");
          if (percent != null) setLineSpacingPercent(paragraphs, percent);
        }
      }
    }
    if (JSON.stringify(target) === JSON.stringify(before)) return;
    const after = structuredClone(target);
    const restore = (snap: typeof before): void => {
      for (const key of Object.keys(target) as (keyof typeof target)[]) delete target[key];
      Object.assign(target, snap);
    };
    this.#pushEdit({
      undo: () => {
        restore(before);
        this.#reproject(sel!.slide);
      },
      redo: () => {
        restore(after);
        this.#reproject(sel!.slide);
      },
    });
    this.#reproject(sel!.slide);
  }

  // ── Hyperlink ────────────────────────────────────────────────────────────

  /** The shared link dialog over the text under edit (the cell a table edit
   *  was in, else the selected shape's body). PowerPoint applies the address
   *  to the object's whole text. */
  #openLinkDialog(): void {
    const runs = this.#runsUnderEdit();
    if (!runs) return;
    const existing = runs.map((run) => run.hyperlink?.url).find((url) => url != null);
    (
      this.shadowRoot?.querySelector("docen-link-dialog") as {
        show(values?: Partial<LinkValues>): void;
      } | null
    )?.show({ href: existing });
  }

  readonly #onLinkOk = (event: Event): void => {
    // The same target rule as the format commands; an in-flight edit writes
    // back first, or the exit-time write-back would clobber the link.
    const at = this.#tableEdit;
    this.#exitTextEditing(true);
    const sel = this.#selection;
    const cell = at ? this.#editedCell(at) : null;
    const body = cell ? null : this.#selectedShapeBody();
    if (!cell && !body) return;
    const target = cell ?? body!;
    const runs = runsIn(cell ? cellParagraphsOf(cell) : bodyParagraphsOf(body!));
    const href = ((event as CustomEvent<LinkValues>).detail?.href ?? "").trim();
    // The dialog's placeholder is not an address; an empty one removes.
    const address = !href || href === "https://" ? undefined : href;
    const before = structuredClone(target);
    for (const run of runs) {
      if (address) run.hyperlink = { url: address };
      else delete run.hyperlink;
    }
    if (JSON.stringify(target) === JSON.stringify(before)) return;
    const after = structuredClone(target);
    const restore = (snap: typeof before): void => {
      for (const key of Object.keys(target) as (keyof typeof target)[]) delete target[key];
      Object.assign(target, snap);
    };
    this.#pushEdit({
      undo: () => {
        restore(before);
        this.#reproject(sel!.slide);
      },
      redo: () => {
        restore(after);
        this.#reproject(sel!.slide);
      },
    });
    this.#reproject(sel!.slide);
  };

  /** The live text runs under edit — the table cell's paragraphs or the
   *  selected shape's body. Null when there is no text target. */
  #runsUnderEdit(at = this.#tableEdit): TextRunOptions[] | null {
    if (at) {
      const cell = this.#editedCell(at);
      return cell ? runsIn(cellParagraphsOf(cell)) : null;
    }
    const body = this.#selectedShapeBody();
    return body ? runsIn(bodyParagraphsOf(body)) : null;
  }

  /** The selected shape's text body, or null (groups' members are a later
   *  batch — same rule as the format commands). */
  #selectedShapeBody(): TextBodyOptions | null {
    const child = this.#selectedChild();
    if (!child || !("shape" in child) || !child.shape.textBody) return null;
    return child.shape.textBody;
  }

  /** Stamp the header's undo/redo buttons from the edit-stack position. */
  #syncQat(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const undo = root.querySelector<HTMLElement>("docen-title-bar [event='undo']");
    const redo = root.querySelector<HTMLElement>("docen-title-bar [event='redo']");
    if (!undo || !redo) return;
    const canUndo = this.#editIndex >= 0;
    const canRedo = this.#editIndex < this.#edits.length - 1;
    undo.toggleAttribute("disabled", !canUndo);
    redo.toggleAttribute("disabled", !canRedo);
  }

  /** Generate the edited deck and hand it to the browser as a download. */
  async #saveAs(): Promise<void> {
    if (!this.#presJson) return;
    const blob = await generatePresentation(this.#presJson, { type: "blob" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = this.getAttribute("filename") ?? "presentation.pptx";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  /** The pointer position as slide-local coordinates — which slide, x/y in
   *  slide px — or null when the pointer is outside the deck strip. */
  #stagePointOf(event: PointerEvent | MouseEvent): { slide: number; x: number; y: number } | null {
    const pres = this.#pres;
    if (!pres) return null;
    const rect = this.#stage.getBoundingClientRect();
    if (rect.width === 0) return null;
    const x = ((event.clientX - rect.left) / rect.width) * pres.widthPx;
    const stripY =
      ((event.clientY - rect.top) / rect.height) *
      (pres.slides.length * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX);
    const slide = Math.floor(stripY / (pres.heightPx + SLIDE_GAP_PX));
    if (slide < 0 || slide >= pres.slides.length) return null;
    return { slide, x, y: stripY - slide * (pres.heightPx + SLIDE_GAP_PX) - SLIDE_GAP_PX };
  }

  readonly #onStagePointerDown = (event: PointerEvent): void => {
    // A click advances the show (PowerPoint's rule) — no gestures inside.
    if (this.#presenting) {
      if (event.button === 0) this.#gotoShowSlide(this.#showingSlide + 1);
      return;
    }
    const presJson = this.#presJson;
    if (!presJson || event.button !== 0) return;
    const point = this.#stagePointOf(event);
    if (!point) return;
    const child = hitSlide(slideHits(presJson.slides?.[point.slide] ?? {}), point.x, point.y);
    // A press outside the object under edit leaves the edit (Word's rule);
    // one inside just moves the caret — the textarea keeps the session. In a
    // table edit a press on another cell of the same table hops there.
    if (this.#textEditor) {
      if (this.#selection?.slide !== point.slide || this.#selection.child !== child)
        this.#select(null);
      else if (this.#tableEdit) {
        // The hop re-focuses the cell editor; the pointerdown's default focus
        // move runs after this handler and would steal the caret back.
        event.preventDefault();
        this.#moveTableCellEditing(point.x, point.y);
      }
      return;
    }
    if (child < 0) return this.#select(null);
    const target = presJson.slides?.[point.slide]?.children?.[child];
    const sameObject = this.#selection?.slide === point.slide && this.#selection.child === child;
    // The selected group: a press on a member descends into it (PowerPoint's
    // second click); a member press selects or drags, a press on empty group
    // canvas falls back to the group itself.
    if (sameObject && target && "group" in target) {
      const hit = memberAt(target, point.x, point.y);
      const member = this.#selection?.member;
      if (member) {
        if (hit && pathsEqual(hit.path, member))
          return this.#overlay?.beginMove(event.clientX, event.clientY);
        if (hit) return this.#select({ slide: point.slide, child, member: hit.path });
        return this.#select({ slide: point.slide, child });
      }
      if (hit) return this.#select({ slide: point.slide, child, member: hit.path });
      return this.#overlay?.beginMove(event.clientX, event.clientY);
    }
    if (sameObject) return this.#overlay?.beginMove(event.clientX, event.clientY);
    this.#select({ slide: point.slide, child });
  };

  readonly #onStageDblClick = (event: MouseEvent): void => {
    if (this.#textEditor) return;
    const point = this.#presJson ? this.#stagePointOf(event) : null;
    // A table under the pointer opens its cell edit; anything else opens the
    // shape text edit (the first click of the double-click selected it).
    if (point) {
      const child = hitSlide(
        slideHits(this.#presJson!.slides?.[point.slide] ?? {}),
        point.x,
        point.y,
      );
      const target =
        child >= 0 ? this.#presJson!.slides?.[point.slide]?.children?.[child] : undefined;
      if (child >= 0 && target && "table" in target)
        return this.#enterTableCellEditing(point.x, point.y);
      // A group's table member opens its cell edit — the first click of the
      // double-click already descended into the member.
      if (child >= 0 && target && "group" in target) {
        const hit = memberAt(target, point.x, point.y);
        const leaf = hit && this.#childAt(target, hit.path);
        if (hit && leaf && "table" in leaf) return this.#enterTableCellEditing(point.x, point.y);
      }
    }
    this.#enterTextEditing();
  };

  readonly #onKeyDown = (event: KeyboardEvent): void => {
    // The show owns the keyboard while presenting: arrow/space paging,
    // Escape leaving. Everything else falls through.
    if (this.#presenting) {
      if (event.key === "Escape") {
        event.preventDefault();
        this.#endShow();
      } else if (["ArrowRight", "ArrowDown", "PageDown", " "].includes(event.key)) {
        event.preventDefault();
        this.#gotoShowSlide(this.#showingSlide + 1);
      } else if (["ArrowLeft", "ArrowUp", "PageUp"].includes(event.key)) {
        event.preventDefault();
        this.#gotoShowSlide(this.#showingSlide - 1);
      } else if (event.key === "Home") {
        this.#gotoShowSlide(0);
      } else if (event.key === "End") {
        this.#gotoShowSlide((this.#pres?.slides.length ?? 1) - 1);
      }
      return;
    }
    // Inside the text edit the textarea owns every key: Enter/Delete type,
    // its Escape handler commits and leaves; undo is the browser's (the
    // textarea's own typing history).
    if (this.#textEditor) return;
    if (event.key === "Escape" && this.#selection) {
      // A member selection climbs back to its group before leaving it.
      if (this.#selection.member) {
        this.#selection = { slide: this.#selection.slide, child: this.#selection.child };
        const hit = this.#selectedBox();
        if (hit) this.#overlay?.show(hit.box, hit.rotation);
        else this.#overlay?.hide();
        return;
      }
      return this.#select(null);
    }
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "z") {
      event.preventDefault();
      return this.#undo();
    }
    if (
      (event.ctrlKey || event.metaKey) &&
      (event.key.toLowerCase() === "y" || (event.shiftKey && event.key.toLowerCase() === "z"))
    ) {
      event.preventDefault();
      return this.#redo();
    }
    if (
      (event.key === "Delete" || event.key === "Backspace") &&
      this.#selection &&
      // A member Delete would splice inside the group — not wired yet, and
      // falling through would remove the whole group under the user.
      !this.#selection.member
    ) {
      event.preventDefault();
      this.#deleteSelected();
    }
  };

  // ── Thumbnails panel ─────────────────────────────────────────────────────

  /** The rail shows the whole deck on one scaled Leafer surface — same paint
   *  loop as the main strip, shrunk through the tree's scale. One canvas
   *  keeps the panel cheap; the selection frame is a DOM sibling that moves
   *  by transform. A `slide` argument swaps just that thumbnail; without it
   *  the surface rebuilds (deck swap, structural slide changes). */
  #renderThumbnails(slide?: number): void {
    const pres = this.#pres;
    const strip = this.thumbStrip;
    const stage = this.shadowRoot?.querySelector<HTMLDivElement>(".thumb-stage");
    const panel = this.shadowRoot?.querySelector<HTMLElement>(".slides-panel");
    if (!pres || !strip || !stage || !panel) return;
    panel.removeAttribute("hidden");
    const scale = THUMB_WIDTH_PX / pres.widthPx;
    this.#thumbScale = scale;
    // Screen-space gaps stay outside the scale: the pitch passes the thumbnail
    // gap divided by the scale so spacing survives the coordinate shrink.
    const pitch = pres.heightPx + THUMB_GAP_PX / scale;
    if (
      slide !== undefined &&
      this.#thumbApp &&
      (this.#thumbApp.tree as unknown as IGroup).children.length === pres.slides.length
    ) {
      repaintSlideAt(
        this.#thumbApp.tree as unknown as IGroup,
        pres,
        slide,
        () => this.#thumbApp?.forceRender(),
        pitch,
        0,
      );
      return;
    }
    this.#thumbApp?.destroy();
    stage.replaceChildren();
    const thumbHeight = pres.heightPx * scale;
    stage.style.width = `${THUMB_WIDTH_PX}px`;
    stage.style.height = `${pres.slides.length * thumbHeight + (pres.slides.length - 1) * THUMB_GAP_PX}px`;
    const app = new App({
      view: stage,
      fill: "transparent",
      tree: { type: "design" },
      move: { disabled: true },
      wheel: { disabled: true },
    });
    this.#thumbApp = app;
    const tree = app.tree as unknown as IGroup;
    tree.scale = { x: scale, y: scale };
    paintSlideDeck(tree, pres, () => app.forceRender(), pitch, 0);
  }

  /** Click a thumbnail → bring that slide to the top of the viewport. */
  readonly #onThumbClick = (event: MouseEvent): void => {
    const pres = this.#pres;
    const strip = this.thumbStrip;
    if (!pres || !strip) return;
    const rect = strip.getBoundingClientRect();
    const pitch = pres.heightPx * this.#thumbScale + THUMB_GAP_PX;
    const index = Math.floor((event.clientY - rect.top) / pitch);
    this.#revealSlide(Math.max(0, Math.min(pres.slides.length - 1, index)));
  };

  #revealSlide(index: number): void {
    const pres = this.#pres;
    const area = this.#area();
    if (!pres || !area) return;
    const pitch = (pres.heightPx + SLIDE_GAP_PX) * (this.#zoom / 100);
    area.scrollTo({ top: index * pitch, behavior: "smooth" });
  }

  /** Move the panel's selection frame onto the active slide. */
  #syncThumbSelection(index: number): void {
    const pres = this.#pres;
    const frame = this.thumbSelection;
    if (!pres || !frame) return;
    const thumbHeight = pres.heightPx * this.#thumbScale;
    frame.style.height = `${thumbHeight}px`;
    frame.style.transform = `translateY(${index * (thumbHeight + THUMB_GAP_PX)}px)`;
  }

  #pickFile(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#file-input")?.click();
  }

  /** File-input open path: stamp the filename attribute so the header menu
   *  mirrors the open deck, then ride the same openPresentation pipeline. */
  readonly #onFileChange = async (event: Event): Promise<void> => {
    const file = (event.target as HTMLInputElement).files?.[0];
    (event.target as HTMLInputElement).value = "";
    if (!file) return;
    this.setAttribute("filename", file.name);
    this.#renderChrome();
    await this.openPresentation(await file.arrayBuffer());
  };

  #syncLang(): void {
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    const lang = this.lang;
    if (lang) workspace?.setAttribute("lang", lang);
    else workspace?.removeAttribute("lang");
    notifyLocaleChange();
  }
}

export default DocenPresentation;
