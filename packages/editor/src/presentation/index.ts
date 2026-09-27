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

import { measureEmu, solidFillOf } from "@docen/core/geometry";
import {
  browserFontMetrics,
  EMU_PER_PX,
  familyOfSlot,
  leaferBaselinePadPx,
  ptToPx,
  type LayoutBlock,
  type LayoutDrawingFill,
  type LayoutDrawingMember,
} from "@docen/layout";
import {
  generatePresentation,
  parsePresentation,
  projectPresentation,
  tableGridOf,
  memberAt,
  memberByPath,
  offsetMemberByPath,
  resizeMemberByPath,
  textBlocks,
  type PresentationOptions,
  type ProjectedPresentation,
  type SlideOptions,
  type SlideChild,
  type TableOptions,
  type TableCellOptions,
  type TransitionOptions,
  type TransitionType,
  type SlideAnimation,
} from "@docen/pptx";
import { customElement, observable } from "@microsoft/fast-element";
import type { DataType } from "@office-open/core";
import type {
  RunFont,
  ShapeType,
  TextBodyOptions,
  TextFont,
  TextRunOptions,
} from "@office-open/core/drawing";
// Leafer ships animate() as a stub that only logs — the show's entrance
// tweens need the real plugin registered.
import "@leafer-in/animate";
import { App, type IGroup } from "leafer-ui";

import { renderRibbonFromSchema } from "../document/ribbon";
import type { Box } from "../drawing/geometry";
import { DrawingOverlay } from "../drawing/overlay";
import { ShapeDrawer, type ShapeDrawRect } from "../drawing/shape-drawer";
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
  nonVisualOf,
  duplicateSlideAt,
  firstCellRunSizeOf,
  insertSlideAt,
  makePicture,
  makeShape,
  makeTable,
  makeTextBox,
  makeFieldBox,
  makeLine,
  makeMediaFrame,
  makeSmartArt,
  reorderChild,
  runsIn,
  setLineSpacingPercent,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  slideNotesOf,
  toggleBullet,
  toggleNumbering,
  toggleRunFlag,
  toggleRunStyle,
  writeSlideNotes,
} from "./commands";
import { DeckHistory } from "./deck-history";
import {
  captureGeometry,
  hitSlide,
  offsetChild,
  rotationOf,
  resizeChild,
  restoreGeometry,
  rotateChild,
  slideHits,
} from "./hit-test";
import { MediaPlayer, type MediaPlayback } from "./media-player";
import { presentationRibbonTabs } from "./ribbon";
import { changedSlides } from "./slide-diff";
import {
  paintSlideDeck,
  SLIDE_GAP_PX,
  repaintSlide,
  THUMB_GAP_PX,
  THUMB_WIDTH_PX,
} from "./slide-paint";
import {
  disposeTextAreaMirror,
  measureTextAreaContent,
  slideFillStyle,
  textBoxFillStyle,
} from "./text-overlay";
import { textOf, writeText, type TextEditSource } from "./text-session";
// Side-effect: register the presentation translation tables.
import "./i18n";

const ZOOM_MIN = 10;
const ZOOM_MAX = 500;

/** The drawing grid pitch: PowerPoint's 0.5" grid in slide px (96/inch). */
const GRID_PITCH_PX = 48;

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

/** The Animate split's presets: the AnimationType tokens the show playback
 *  maps onto Leafer's own tweening (appear = instant, no tween). */
const ANIMATION_PRESETS: ReadonlySet<string> = new Set(["appear", "fade", "fly", "zoom"]);

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

/** A media file extension → its source-model family and type. */
const MEDIA_FILE_TYPES: ReadonlyMap<
  string,
  { media: "video" | "audio"; type: "mp4" | "mov" | "wmv" | "avi" | "mp3" | "wav" | "wma" | "aac" }
> = new Map(
  (
    [
      ["mp4", "video", "mp4"],
      ["mov", "video", "mov"],
      ["wmv", "video", "wmv"],
      ["avi", "video", "avi"],
      ["mp3", "audio", "mp3"],
      ["wav", "audio", "wav"],
      ["wma", "audio", "wma"],
      ["aac", "audio", "aac"],
    ] as const
  ).map(([extension, media, type]) => [extension, { media, type }]),
);

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
  "video",
  "audio",
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
  "animate",
  "add-animation",
  "slide-number",
  "date-time",
  "select",
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
  /** The projection backing the current main-strip nodes; member diffs pair
   *  old and new batches against it. */
  #renderedMainPres: ProjectedPresentation | null = null;
  /** The thumbnail projection backing its current nodes; thumbnail repaints
   *  use it to pair source-child batches just like the main strip. */
  #renderedThumbPres: ProjectedPresentation | null = null;
  #app: App | null = null;
  #thumbApp: App | null = null;
  #thumbScale = 1;
  #zoom = 100;
  #overlay: DrawingOverlay | null = null;
  #shapeDrawer: ShapeDrawer | null = null;
  #mediaPlayer: MediaPlayer | null = null;
  /** The selected object: the slide child, or — with `member` set — the
   *  group-nested member it descended into (child indexes under the group). */
  #selection: { slide: number; child: number; member?: number[] } | null = null;
  /** The in-place shape-text editor: the textarea floats over the shape while
   *  it holds the session; every keystroke writes through into the shape, so
   *  the text lives in the model as it is typed. */
  #textEditor: HTMLTextAreaElement | null = null;
  /** The edited shape's text body as it entered the session — the undo
   *  baseline the exit compares against. */
  #textEditBefore: TextBodyOptions | null = null;
  /** The pending write-through frame (canceled by a commit-path exit). */
  #textEditRaf = 0;
  /** The table cell under the in-place edit (grid row/col — the session
   *  itself lives in #textEditor; null there means a shape text edit). */
  #tableEdit: { row: number; col: number } | null = null;
  /** Drawing gridlines visibility (a view state — not part of the deck). */
  #gridlines = false;
  readonly #history = new DeckHistory();
  /** Changed projected slides since the previous paint; undefined forces a
   *  full paint. The next render consumes and clears it. */
  #changedSlides?: Set<number>;
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
  /** The selection pane's object list (chrome.ts). */
  @observable selectList?: HTMLElement;

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
      .querySelector<HTMLInputElement>("#media-input")
      ?.addEventListener("change", this.#onMediaChange as EventListener);
    root
      .querySelector("docen-find-replace-dialog")
      ?.addEventListener("find-replace:action", this.#onFindReplace as EventListener);
    root
      .querySelector("docen-link-dialog")
      ?.addEventListener("link:ok", this.#onLinkOk as EventListener);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.selectList?.addEventListener("click", this.#onSelectListClick);
    this.#area()?.addEventListener("wheel", this.#onShowWheel, { passive: false });
    this.#area()?.addEventListener("scroll", this.#onScroll);
    this.thumbStrip?.addEventListener("click", this.#onThumbClick);
    // Capture: Leafer's own app view can consume a pointer before bubbling;
    // the editor must get first refusal for an armed drag-to-draw tool.
    this.#stage.addEventListener("pointerdown", this.#onStagePointerDown, true);
    this.#stage.addEventListener("dblclick", this.#onStageDblClick);
    document.addEventListener("keydown", this.#onKeyDown);
    // The selection frame lives over the slide surface (the canvas element
    // Leafer mounts is recreated per deck render — the overlay survives it).
    this.#overlay = new DrawingOverlay({
      scale: () => this.#zoom / 100,
      applyBox: (box) => {
        const selection = this.#selection;
        const member = selection?.member;
        const group = member
          ? this.#presJson?.slides?.[selection.slide]?.children?.[selection.child]
          : undefined;
        this.#applyGesture(() =>
          member && group
            ? resizeMemberByPath(group, member, this.#slideBoxOf(box))
            : resizeChild(this.#selectedChild()!, this.#slideBoxOf(box)),
        );
      },
      applyOffset: (dx, dy) => {
        // The walk folds ancestor scales and spins into one inverse mapping.
        const selection = this.#selection;
        const member = selection?.member;
        const group = member
          ? this.#presJson?.slides?.[selection.slide]?.children?.[selection.child]
          : undefined;
        this.#applyGesture(() => {
          if (member && group) offsetMemberByPath(group, member, dx, dy);
          else offsetChild(this.#selectedChild()!, dx, dy);
        });
      },
      applyRotation: (delta) => {
        this.#applyGesture((child) => rotateChild(child, delta));
      },
    });
    this.#canvasHost().append(this.#overlay.el);
    this.#mediaPlayer = new MediaPlayer(this.#canvasHost());
    // Shapes/Text Box arm this drag-to-draw tool. The host maps the pointer
    // to the pinned slide frame, so a sweep can start and land on any slide.
    this.#shapeDrawer = new ShapeDrawer(
      {
        frameAt: (clientX, clientY) => {
          const pres = this.#pres;
          const point = this.#stagePointOf({ clientX, clientY } as PointerEvent);
          if (!pres || !point) return null;
          const surfaceRect = this.#canvasHost().getBoundingClientRect();
          const stageRect = this.#stage.getBoundingClientRect();
          if (stageRect.width === 0 || stageRect.height === 0) return null;
          const scale = stageRect.width / pres.widthPx;
          return {
            slide: point.slide,
            left: stageRect.left - surfaceRect.left,
            top:
              stageRect.top -
              surfaceRect.top +
              (point.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX) * scale,
            width: pres.widthPx,
            height: pres.heightPx,
            scale,
          };
        },
        apply: (preset, rect) => this.#insertDrawnShape(preset, rect),
      },
      this.#canvasHost(),
    );
    // Locale switches re-stamp the chrome (header + ribbon labels).
    this.#unsubLang = observeLang(() => this.#renderChrome());
    this.#renderChrome();
    if (this.#pres) this.#renderDeck();
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#shapeDrawer?.destroy();
    this.#shapeDrawer = null;
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
      ?.querySelector<HTMLInputElement>("#media-input")
      ?.removeEventListener("change", this.#onMediaChange as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#onFindReplace as EventListener);
    this.shadowRoot
      ?.querySelector("docen-link-dialog")
      ?.removeEventListener("link:ok", this.#onLinkOk as EventListener);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.selectList?.removeEventListener("click", this.#onSelectListClick);
    this.#area()?.removeEventListener("wheel", this.#onShowWheel);
    this.#area()?.removeEventListener("scroll", this.#onScroll);
    this.thumbStrip?.removeEventListener("click", this.#onThumbClick);
    this.#stage.removeEventListener("pointerdown", this.#onStagePointerDown);
    this.#stage.removeEventListener("dblclick", this.#onStageDblClick);
    document.removeEventListener("keydown", this.#onKeyDown);
    this.#overlay?.el.remove();
    this.#overlay = null;
    this.#mediaPlayer?.destroy();
    this.#mediaPlayer = null;
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
    this.#setPresentation(pres);
  }

  /** DocenHost surface: the editable deck is plain, structured JSON. A copy
   *  keeps host-side mutation from bypassing the editor's history/rendering. */
  getContent(): PresentationOptions | null {
    return this.#presJson ? structuredClone(this.#presJson) : null;
  }

  setContent(content: unknown): void {
    if (!content || typeof content !== "object" || Array.isArray(content)) return;
    this.#setPresentation(structuredClone(content) as PresentationOptions);
  }

  /** Generate the current model for a host that owns persistence. */
  async savePresentation(): Promise<Uint8Array> {
    if (!this.#presJson) throw new Error("No presentation is open");
    return generatePresentation(this.#presJson, { type: "uint8array" });
  }

  /** Drop the open deck and reset the chrome. */
  closePresentation(): void {
    this.#presJson = null;
    this.#pres = null;
    this.#exitTextEditing(false);
    this.#selection = null;
    this.#history.clear();
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

  /** Install a new deck and reset all per-deck edit state. */
  #setPresentation(pres: PresentationOptions): void {
    this.#presJson = pres;
    this.#exitTextEditing(false);
    this.#selection = null;
    this.#history.clear();
    this.#notesSlideIndex = 0;
    this.#syncNotesPane(0);
    this.#pres = projectPresentation(pres);
    this.#changedSlides = undefined;
    // The base class attaches an empty shadowRoot ahead of connect, so
    // shadowRoot presence does not mean the template stamped — same bail
    // #renderChrome does until connectedCallback renders the stored deck.
    if (!this.shadowRoot?.querySelector(".stage")) return;
    this.#renderDeck();
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
      { id: "undo", icon: "undo", disabled: !this.#history.canUndo },
      { id: "redo", icon: "redo", disabled: !this.#history.canRedo },
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
    else if (name === "shapes") this.#armShape(event.detail?.value);
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
    else if (name === "animate" || name === "add-animation") {
      this.#applyAnimation(event.detail?.value);
    } else if (name === "slide-number" || name === "date-time") this.#insertFieldBox(name);
    else if (name === "smartart") this.#insertSmartArt();
    else if (name === "video" || name === "audio") this.#pickMedia();
    else if (name === "select") this.#toggleSelectionPane();
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
    // A scrolled strip strands the media overlay (its screen box is a
    // snapshot) — playback yields to the scroll, like PowerPoint's editor.
    this.#mediaPlayer?.hide();
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
    this.#renderSelectionPane();
  }

  /** Public zoom surface aligned with `<docen-document>`: callers don't need
   *  to know the ribbon event or status-bar attribute behind the control. */
  setZoom(pct: number): void {
    this.#setZoom(pct);
  }

  getZoom(): number {
    return this.#zoom;
  }

  /** Repaint the deck. With `slide` set (a single-slide edit), only that
   *  slide's group swaps out — the rest of the strip stays untouched; without
   *  it the whole tree repaints (open, structural slide changes). */
  #renderDeck(slide?: number): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#mediaPlayer?.hide();
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
    const slideChanged = slide === undefined || this.#consumeSlideChange(slide);
    if (slideChanged && slide !== undefined && tree.children.length === pres.slides.length) {
      repaintSlide(
        tree,
        pres,
        slide,
        () => this.#app?.forceRender(),
        SLIDE_GAP_PX + slide * (pres.heightPx + SLIDE_GAP_PX),
        this.#renderedMainPres?.slides[slide],
      );
    } else {
      tree.clear();
      paintSlideDeck(tree, pres, () => this.#app?.forceRender());
    }
    this.#applyZoom();
    this.#renderThumbnails(slideChanged ? slide : undefined);
    this.#syncSlideIndicator();
    if (slideChanged || slide === undefined) this.#renderedMainPres = pres;
  }

  /** Whether slide `index` needs a repaint; first paints consume an undefined
   *  diff as “all changed”. */
  #consumeSlideChange(index: number): boolean {
    const changed = this.#changedSlides;
    return !changed || changed.has(index);
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
        rotation: m.rotation,
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

  /** Select a slide object (or clear). The frame shows the strip-space box. */
  #select(sel: { slide: number; child: number; member?: number[] } | null): void {
    if (this.#textEditor) this.#exitTextEditing(true);
    this.#selection = sel;
    if (!sel) this.#overlay?.hide();
    else {
      const hit = this.#selectedBox();
      if (hit) this.#overlay?.show(hit.box, hit.rotation);
      else this.#overlay?.hide();
    }
    this.#renderSelectionPane();
    this.#syncTextFormatControls();
  }

  /** Double-click on a shape: float a textarea over its text area and hand
   *  the session to it (PowerPoint's in-place text edit). The overlay's
   *  resize handles step aside for the duration. The overlay styles from
   *  the projected textBox member the canvas painted (insets, first-run
   *  face, the shape's fill); an empty body projects no member, so the
   *  fallback re-derives the same values through textBlocks — the
   *  projection's own defaults, never the overlay's assumptions. Every
   *  keystroke writes through into the shape (raf-merged), so the canvas
   *  repaints the typed text under the opaque overlay and the model never
   *  drifts from what the user sees. */
  #enterTextEditing(): void {
    const sel = this.#selection;
    const pres = this.#pres;
    const presJson = this.#presJson;
    const child = this.#selectedChild();
    const source: TextEditSource | null =
      child && "shape" in child ? { kind: "shape", child: child.shape } : null;
    const text = source ? textOf(source) : null;
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
    // equality identifies it. An empty body projects none — the fallbacks
    // below resolve the same styles from the child itself.
    const member = pres.slides[sel.slide]?.members.find(
      (m): m is Extract<LayoutDrawingMember, { kind: "textBox" }> => {
        if (m.kind !== "textBox") return false;
        return (
          m.x === box!.x && m.y === box!.y && m.width === box!.width && m.height === box!.height
        );
      },
    );
    const body = "shape" in child ? child.shape.textBody : undefined;
    const bp = body?.bodyProperties;
    const anchorOf = (value: string | undefined): "top" | "center" | "bottom" =>
      value === "center" || value === "bottom" ? value : "top";
    // a:bodyPr's DrawingML default insets (0.1" sides, 0.05" top/bottom) —
    // the same defaults the projection's textBox member applies.
    const insOf = (v: unknown, def: number): number => (measureEmu(v) ?? def) / EMU_PER_PX;
    const ins = member
      ? {
          left: member.insets?.left ?? 0,
          top: member.insets?.top ?? 0,
          right: member.insets?.right ?? 0,
          bottom: member.insets?.bottom ?? 0,
        }
      : {
          left: insOf(bp?.lIns, 91440),
          top: insOf(bp?.tIns, 45720),
          right: insOf(bp?.rIns, 91440),
          bottom: insOf(bp?.bIns, 45720),
        };
    const anchor = member?.anchor ?? anchorOf(bp?.anchor ?? body?.anchor);
    const autoFit = member?.autoFit === true;
    const fill: LayoutDrawingFill | undefined =
      member?.fill ?? solidFillOf("shape" in child ? child.shape.properties?.fill : undefined);
    // The first paragraph's text-run style — the member's blocks when the
    // projection made one, else textBlocks' resolution of the same body (its
    // defaults ARE the painted defaults). The bare constant covers a body
    // with no paragraphs at all.
    const blocks =
      member?.blocks ??
      (body
        ? textBlocks(body, {
            slideNumber: sel.slide + 1,
            slideCount: this.#presJson?.slides?.length ?? 0,
            now: new Date(),
          })
        : []);
    const first = blocks[0]?.kind === "paragraph" ? blocks[0] : undefined;
    const run = first?.inline.find((i) => i.kind === "text")?.style ??
      first?.defaultTextStyle ?? { family: "Calibri", sizePx: ptToPx(18) };
    // The painted stack advances lines by the engine's normal line height and
    // hangs shape-text glyphs at 0.85 em below the line top (the painter's
    // shape-text model). The overlay reproduces both: an explicit line-height
    // sized from the same metric source the paint context uses, plus a
    // baseline correction that moves the CSS baseline (half-leading + font
    // ascent) onto the painted depth — CSS offers no direct baseline control.
    const face = familyOfSlot(run.family, false);
    const rotation = this.#selectedBox()?.rotation ?? rotationOf(child);
    const linePx =
      browserFontMetrics.normalRatio({
        family: face,
        bold: run.bold === true,
        italic: run.italic === true,
      }) * run.sizePx;
    const baselineShift = this.#baselineShiftOf(
      face,
      run.bold === true,
      run.italic === true,
      run.sizePx,
      linePx,
      scale,
    );
    const editor = document.createElement("textarea");
    editor.className = "shape-text-editor";
    // The rows=2 default inflates scrollHeight to two lines — the slack math
    // (and every grow pass) must measure the real content.
    editor.rows = 1;
    editor.value = text;
    // Opaque: while the session is open this overlay IS the text surface —
    // translucency would ghost the canvas painting through it.
    const slide = pres.slides[sel.slide]!;
    const fillStyle = fill
      ? textBoxFillStyle(fill, member?.opacity)
      : slideFillStyle(slide.background, pres.widthPx, pres.heightPx);
    Object.assign(editor.style, {
      left: `${box.x * scale}px`,
      top: `${stripY * scale}px`,
      width: `${box.width * scale}px`,
      height: `${box.height * scale}px`,
      padding: `${ins.top * scale}px ${ins.right * scale}px ${ins.bottom * scale}px ${ins.left * scale}px`,
      fontFamily: JSON.stringify(run.family),
      fontSize: `${run.sizePx * scale}px`,
      lineHeight: `${linePx * scale}px`,
      textAlign: TEXT_ALIGN_OF[first?.align ?? "left"] ?? "left",
      color: run.color ? `#${run.color}` : TEXT_INK,
      transformOrigin: `${(box.width * scale) / 2}px ${(box.height * scale) / 2}px`,
      ...(rotation ? { transform: `rotate(${rotation}deg)` } : {}),
      ...(run.bold ? { fontWeight: "bold" } : {}),
      ...(run.italic ? { fontStyle: "italic" } : {}),
    });
    Object.assign(editor.style, fillStyle);
    // Keep the text visible as it grows, and keep the box recognizable as it
    // doesn't: the edit frame starts at the shape's full height (a short text
    // must not shrink the fill/border rectangle under the user — PowerPoint
    // keeps the frame too) and only grows past it when the text overflows.
    // The baseline correction rides the top inset in every branch; top-anchored
    // boxes grow the frame (height only, the width is the shape's), center/
    // bottom re-run the painter's slack math — the stack starts at the top
    // inset plus half (or all) of the leftover inner height, and overflow
    // spills below the box like the painted stack does.
    const syncLayout = (): void => {
      editor.style.height = "auto";
      editor.style.paddingTop = `${ins.top * scale + baselineShift}px`;
      if (anchor === "top") {
        const content = measureTextAreaContent(editor);
        editor.style.height = `${Math.max(box.height * scale, content + (ins.top + ins.bottom) * scale)}px`;
        editor.scrollTop = 0;
        return;
      }
      const padTB = (ins.top + ins.bottom) * scale;
      const content = measureTextAreaContent(editor);
      const slack = autoFit ? 0 : box.height * scale - padTB - content;
      if (slack >= 0) {
        editor.style.height = `${box.height * scale}px`;
        editor.style.paddingTop = `${ins.top * scale + baselineShift + (anchor === "center" ? slack / 2 : slack)}px`;
      } else {
        editor.style.height = `${content + padTB}px`;
      }
      editor.scrollTop = 0;
    };
    // The write-through: the typed text lands in the shape (raf-merged) and
    // the slide reprojects — the canvas under the overlay paints exactly the
    // session's text. IME composition waits for the commit event (partial
    // pinyin must not hit the model mid-session).
    let composing = false;
    let lastWritten = text;
    const writeThrough = (): void => {
      if (composing || this.#textEditRaf) return;
      this.#textEditRaf = requestAnimationFrame(() => {
        this.#textEditRaf = 0;
        if (this.#textEditor !== editor || editor.value === lastWritten) return;
        lastWritten = editor.value;
        writeText(source!, editor.value);
        this.#reproject(sel.slide);
      });
    };
    editor.addEventListener("keydown", (event) => {
      // Escape leaves the edit (committing); typing keys stay in the textarea.
      if (event.key === "Escape") {
        event.stopPropagation();
        this.#exitTextEditing(true);
      }
      event.stopPropagation();
    });
    editor.addEventListener("input", () => {
      syncLayout();
      writeThrough();
    });
    editor.addEventListener("compositionstart", () => {
      composing = true;
    });
    editor.addEventListener("compositionend", () => {
      composing = false;
      syncLayout();
      writeThrough();
    });
    this.#canvasHost().append(editor);
    // Measure with the element in the tree — the first pass sizes a
    // top-anchored frame and shifts center/bottom onto the painted baseline.
    syncLayout();
    editor.focus();
    editor.select();
    this.#textEditor = editor;
    this.#textEditBefore = body ? structuredClone(body) : null;
    this.#overlay?.hide();
    this.#syncTextFormatControls();
  }

  /** Leave the text edit. The session's keystrokes have written through into
   *  the shape, so `write` only records the undo step (the entry body vs
   *  now) and repaints once, synchronously — the pending raf's final state
   *  is flushed first so the recorded "after" is the text the user sees.
   *  Plain exits just drop the overlay. A table cell session dispatches to
   *  its own exit. */
  #exitTextEditing(write: boolean): void {
    if (this.#tableEdit) return this.#exitTableCellEditing(write);
    const editor = this.#textEditor;
    this.#textEditor = null;
    if (this.#textEditRaf) {
      cancelAnimationFrame(this.#textEditRaf);
      this.#textEditRaf = 0;
    }
    editor?.remove();
    disposeTextAreaMirror(editor);
    const sel = this.#selection;
    const child = this.#selectedChild();
    const before = this.#textEditBefore;
    this.#textEditBefore = null;
    if (editor && write && sel && child && "shape" in child && before && child.shape.textBody) {
      const shape = child.shape;
      const source: TextEditSource = { kind: "shape", child: child.shape };
      if (editor.value !== textOf(source)) writeText(source, editor.value);
      const after = structuredClone(shape.textBody);
      if (JSON.stringify(before) !== JSON.stringify(after)) {
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
    this.#restoreOverlay();
    this.#syncTextFormatControls();
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
    const syncLayout = (): void => {
      editor.style.height = "auto";
      editor.style.paddingTop = `${m.top * scale + baselineShift}px`;
      const boxH = rect.height * scale;
      const padTB = (m.top + m.bottom) * scale;
      if (anchor === "top") {
        const topContent = measureTextAreaContent(editor);
        editor.style.height = `${Math.max(boxH, topContent + padTB)}px`;
        editor.scrollTop = 0;
        return;
      }
      const content = measureTextAreaContent(editor);
      const slack = boxH - padTB - content;
      if (slack >= 0) {
        editor.style.height = `${boxH}px`;
        editor.style.paddingTop = `${m.top * scale + baselineShift + (anchor === "center" ? slack / 2 : slack)}px`;
      } else {
        editor.style.height = `${content + padTB}px`;
      }
      editor.scrollTop = 0;
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
    disposeTextAreaMirror(editor);
    const sel = this.#selection;
    const cell = editor && write ? this.#editedCell(at) : null;
    if (cell) {
      const source: TextEditSource = { kind: "cell", cell };
      if (editor!.value !== textOf(source)) {
        const before = structuredClone(cell);
        writeText(source, editor!.value);
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
    this.#syncTextFormatControls();
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
    const previous = this.#pres;
    this.#pres = projectPresentation(this.#presJson);
    this.#changedSlides = changedSlides(previous, this.#pres);
    this.#renderDeck(slide);
    this.#changedSlides = undefined;
    // The paint restarted under the same selection — the frame snaps to the
    // current geometry (and rejects itself if the box vanished).
    const hit = this.#selectedBox();
    if (hit) this.#overlay?.refresh(hit.box, hit.rotation);
    else this.#select(null);
    this.#renderSelectionPane();
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
    this.#history.push(edit);
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
    if (!this.#history.canUndo) return;
    this.#history.undo();
    // An in-flight text edit commits first, so undo revokes it (the freshest
    // step) rather than the step before — the user's intent either way.
    this.#exitTextEditing(true);
    this.#syncQat();
  }

  #redo(): void {
    if (!this.#history.canRedo) return;
    this.#history.redo();
    this.#exitTextEditing(true);
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
  #insertChild(child: SlideChild, slideAt = this.#activeSlideIndex()): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const slide = slideAt;
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

  /** Text Box enters drag mode; the draw lands the body on the exact slide. */
  #insertTextBox(): void {
    this.#shapeDrawer?.arm("text-box");
  }

  /** Commit a completed drawer sweep: normal presets use the drawn frame,
   *  straight tokens use endpoint children so direction survives. */
  #insertDrawnShape(preset: string, rect: ShapeDrawRect): void {
    const pres = this.#pres;
    if (!pres) return;
    const placement = { x: rect.x, y: rect.y, w: rect.w, h: rect.h };
    const child =
      preset === "text-box"
        ? makeTextBox(pres.widthPx, pres.heightPx, "", placement)
        : preset === "line" || preset === "straightConnector1"
          ? makeLine(preset, placement)
          : makeShape(pres.widthPx, pres.heightPx, preset as ShapeType, placement);
    this.#insertChild(child, rect.slide);
  }

  /** The footer field inserts (slide number bottom-right, date-time bottom-
   *  left): an a:fld box whose cached text shows the live value now — the
   *  file's field re-evaluates per slide in a real renderer. */
  #insertFieldBox(name: string): void {
    const pres = this.#pres;
    if (!pres) return;
    const page = this.#activeSlideIndex() + 1;
    const child =
      name === "slide-number"
        ? makeFieldBox(pres.widthPx, pres.heightPx, "slidenum", String(page), "right")
        : makeFieldBox(
            pres.widthPx,
            pres.heightPx,
            "datetimeFigureOut",
            new Date().toLocaleDateString(),
            "left",
          );
    this.#insertChild(child);
  }

  #insertTable(): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeTable(pres.widthPx, pres.heightPx));
  }

  #insertSmartArt(): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeSmartArt(pres.widthPx, pres.heightPx));
  }

  /** The Shapes gallery's pick arms the drawer; the value is the prstGeom
   *  token (the main button draws the default rectangle). */
  #armShape(geometry?: string): void {
    this.#shapeDrawer?.arm(geometry ?? "rect");
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

  /** The Animate split's pick on the selected object: one entrance preset as
   *  a SlideAnimation addressed by the shape's cNvPr name (an unnamed shape
   *  takes one unique among its siblings — the file format resolves names to
   *  spTgt ids at compile time). "none" clears the shape's entry; a pick that
   *  changes nothing records nothing. */
  #applyAnimation(value?: string): void {
    const preset = value ?? "fade";
    if (preset !== "none" && !ANIMATION_PRESETS.has(preset)) return;
    const sel = this.#selection;
    const host = this.#presJson?.slides?.[sel?.slide ?? -1];
    const child = host?.children?.[sel?.child ?? -1];
    const nv = child ? nonVisualOf(child) : null;
    if (!host || !nv) return;
    const taken = new Set(
      (host.children ?? [])
        .map((c) => nonVisualOf(c)?.name?.trim())
        .filter((n): n is string => !!n),
    );
    let name = nv.name?.trim();
    if (!name) {
      let n = (sel?.child ?? 0) + 1;
      while (taken.has((name = `Shape ${n}`))) n++;
      nv.name = name;
    }
    const before = host.animations;
    const entries = Array.isArray(before) ? before : [];
    const kept = entries.filter((e) => typeof e === "object" && e.shapeName !== name);
    const after: SlideAnimation[] | undefined =
      preset === "none"
        ? kept.length > 0
          ? kept
          : undefined
        : [...kept, { type: preset as SlideAnimation["type"], shapeName: name }];
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const restore = (value: SlideAnimation[] | string | undefined): void => {
      if (value === undefined) delete host.animations;
      else host.animations = structuredClone(value);
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
  }

  // ── Slide show ───────────────────────────────────────────────────────────

  /** Enter the presenting state: the chrome steps aside, the zoom fits one
   *  slide to the viewport, and the browser goes fullscreen when allowed (a
   *  denied request only means the window stays windowed). */
  #startShow(from: "beginning" | "current"): void {
    const pres = this.#pres;
    if (!pres || this.#presenting) return;
    this.#exitTextEditing(true);
    // The show runs chrome-free: the selection frame and its handles step
    // aside with everything else.
    this.#select(null);
    this.#presenting = true;
    this.#savedZoom = this.#zoom;
    this.#workspaceEl()?.classList.add("presenting");
    const area = this.#area();
    if (area) {
      const fit = Math.floor(((area.clientHeight - SLIDE_GAP_PX) / pres.heightPx) * 100);
      this.#setZoom(Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, fit)));
    }
    // A full repaint first: the show's animations address elements by their
    // place in the tree's child order, and the partial repaints'
    // remove-and-append has long since shuffled that order.
    this.#renderDeck();
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
   *  paging has no smooth travel — and play its entrance animations. */
  #gotoShowSlide(index: number): void {
    const pres = this.#pres;
    const area = this.#area();
    if (!pres || !area) return;
    this.#mediaPlayer?.hide();
    this.#showingSlide = Math.max(0, Math.min(pres.slides.length - 1, index));
    const pitch = (pres.heightPx + SLIDE_GAP_PX) * (this.#zoom / 100);
    area.scrollTo({ top: this.#showingSlide * pitch, behavior: "instant" as ScrollBehavior });
    this.#syncSlideIndicator();
    this.#playShowAnimations(this.#showingSlide);
  }

  /** Play a slide's entrance animations in the live show: each entry tweens
   *  the elements its shape paints, in listed order with a stagger. Members
   *  collect by containment (a simple member sits at the shape's box, a
   *  rotated one at its center, a text body scatters its lines inside it),
   *  so the paint shape never has to be mirrored here; effects outside the
   *  preset set play as a fade. */
  #playShowAnimations(index: number): void {
    const slide = this.#presJson?.slides?.[index];
    const entries = slide?.animations;
    const slideGroup = (this.#app?.tree as unknown as IGroup | undefined)?.children[index] as
      | IGroup
      | undefined;
    if (!slide || !slideGroup || !Array.isArray(entries) || entries.length === 0) return;
    const hits = slideHits(slide);
    const inBox = (box: Box, x: number, y: number): boolean =>
      x >= box.x - 1 && x <= box.x + box.width + 1 && y >= box.y - 1 && y <= box.y + box.height + 1;
    let delay = 0;
    for (const entry of entries) {
      // Entrances only: a parsed exit/emphasis entry waits for a follow-up
      // batch rather than playing backwards.
      if (entry.class === "exit" || entry.class === "emphasis" || entry.class === "mediaCall")
        continue;
      const hit = entry.shapeName
        ? hits.find((h) => nonVisualOf(slide.children![h.child]!)?.name === entry.shapeName)
        : undefined;
      if (!hit) continue;
      const duration = entry.duration ?? 500;
      const options = { duration, delay, jump: true, easing: "ease-out" as const };
      for (const el of slideGroup.children.filter(
        // An unpainted element carries no position — NaN fails the box test.
        (member, i) => i > 0 && inBox(hit.box, member.x ?? NaN, member.y ?? NaN),
      )) {
        const top = el.y ?? 0;
        if (entry.type === "fly") {
          el.animate(
            [
              { y: top + hit.box.height / 3, opacity: 0 },
              { y: top, opacity: 1 },
            ],
            options,
          );
        } else if (entry.type === "zoom") {
          el.around = "center";
          el.animate(
            [
              { scale: 0.3, opacity: 0 },
              { scale: 1, opacity: 1 },
            ],
            options,
          );
        } else if (entry.type !== "appear") {
          el.animate([{ opacity: 0 }, { opacity: 1 }], options);
        }
      }
      delay += duration;
    }
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

  // ── Selection pane ───────────────────────────────────────────────────────

  #selectionPane(): HTMLElement | null {
    return this.shadowRoot?.querySelector(".select-pane") ?? null;
  }

  /** Show/hide the pane (a view state — the list re-stamps when it shows). */
  #toggleSelectionPane(): void {
    const pane = this.#selectionPane();
    if (!pane) return;
    pane.hidden = !pane.hidden;
    if (!pane.hidden) this.#renderSelectionPane();
  }

  /** The selection pane's row label: the cNvPr name when the child carries
   *  one, else the localized kind plus its position. */
  #childLabel(child: SlideChild, index: number): { label: string; hidden: boolean } {
    const nv = nonVisualOf(child);
    const kind =
      "shape" in child
        ? "shape"
        : "picture" in child
          ? "picture"
          : "line" in child
            ? "line"
            : "connector" in child
              ? "connector"
              : "group" in child
                ? "group"
                : "chart" in child
                  ? "chart"
                  : "smartart" in child
                    ? "smartart"
                    : "video" in child
                      ? "video"
                      : "audio" in child
                        ? "audio"
                        : "table";
    const name = nv?.name?.trim();
    return {
      label: name || `${t(`ppt.select.${kind}`, this)} ${index + 1}`,
      hidden: nv?.hidden === true,
    };
  }

  /** Stamp the active slide's object list. A no-op while the pane is hidden —
   *  every refresh trigger re-runs this when it shows. */
  #renderSelectionPane(): void {
    const pane = this.#selectionPane();
    const list = this.selectList;
    if (!pane || pane.hidden || !list) return;
    const head = pane.querySelector<HTMLElement>(".select-head");
    if (head) head.textContent = t("ppt.select.title", this);
    const slide = this.#activeSlideIndex();
    const children = this.#presJson?.slides?.[slide]?.children ?? [];
    list.innerHTML = children
      .map((child, i) => {
        const { label, hidden } = this.#childLabel(child, i);
        const active =
          this.#selection?.slide === slide &&
          this.#selection.child === i &&
          !this.#selection.member;
        return `<div class="select-row" data-child="${i}" data-hidden="${hidden}"${active ? ' data-active="true"' : ""}><button class="select-eye" data-eye="${i}">${hidden ? "○" : "●"}</button><span class="select-name">${escapeHtml(label)}</span></div>`;
      })
      .join("");
  }

  readonly #onSelectListClick = (event: Event): void => {
    const target = event.target as HTMLElement;
    const eye = target.closest<HTMLElement>("[data-eye]");
    if (eye) return this.#toggleChildHidden(Number(eye.dataset.eye));
    const row = target.closest<HTMLElement>("[data-child]");
    if (!row) return;
    const slide = this.#activeSlideIndex();
    this.#select({ slide, child: Number(row.dataset.child) });
    this.#revealSlide(slide);
  };

  /** The eye toggle: cNvPr @hidden on the child (tables have no surface and
   *  stay put) as one reversible edit — hidden objects drop out of the paint. */
  #toggleChildHidden(index: number): void {
    const slide = this.#activeSlideIndex();
    const child = this.#presJson?.slides?.[slide]?.children?.[index];
    const nv = child ? nonVisualOf(child) : null;
    if (!nv) return;
    const before = nv.hidden === true;
    const apply = (hidden: boolean): void => {
      if (hidden) nv.hidden = true;
      else delete nv.hidden;
      this.#reproject(slide);
      this.#renderSelectionPane();
    };
    this.#pushEdit({ undo: () => apply(before), redo: () => apply(!before) });
    apply(!before);
  }

  #pickPicture(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#picture-input")?.click();
  }

  #pickMedia(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#media-input")?.click();
  }

  /** File-input media path: native bytes enter the source model unchanged;
   *  the projection supplies the stable poster/player canvas. */
  readonly #onMediaChange = async (event: Event): Promise<void> => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    const pres = this.#pres;
    const extension = file?.name.split(".").pop()?.toLowerCase() ?? "";
    const spec = MEDIA_FILE_TYPES.get(extension);
    if (!file || !pres || !spec) return;
    this.#insertChild(
      makeMediaFrame(
        pres.widthPx,
        pres.heightPx,
        spec.media,
        new Uint8Array(await file.arrayBuffer()),
        spec.type,
        file.name,
      ),
    );
  };

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
    this.#syncTextFormatControls();
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
    const canUndo = this.#history.canUndo;
    const canRedo = this.#history.canRedo;
    undo.toggleAttribute("disabled", !canUndo);
    redo.toggleAttribute("disabled", !canRedo);
  }

  /** Stamp the font face/size comboboxes from the text under edit — the
   *  document editor's selection sync at this editor's grain: the common
   *  value across the target's runs, blank when mixed or no target. Runs
   *  without explicit props read as the projection's painted defaults
   *  (Calibri 18pt), so the boxes show what the canvas would draw. */
  #syncTextFormatControls(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const faceOf = (font: RunFont | undefined): string => {
      if (typeof font === "string") return font || "Calibri";
      const face = (f: TextFont | undefined): string | undefined =>
        typeof f === "string" ? f || undefined : f?.typeface;
      return face(font?.latin) ?? face(font?.eastAsia) ?? "Calibri";
    };
    const runs = this.#runsUnderEdit() ?? [];
    const common = (values: string[]): string =>
      values.length > 0 && values.every((v) => v === values[0]) ? values[0] : "";
    const face = common(runs.map((run) => faceOf(run.font)));
    const size = common(runs.map((run) => String(run.size ?? 18)));
    for (const [event, value] of [
      ["font-face", face],
      ["font-size", size],
    ] as const) {
      const box = root.querySelector<HTMLElement>(`docen-ribbon-combobox[event='${event}']`);
      if (box && box.getAttribute("value") !== value) box.setAttribute("value", value);
    }
  }

  /** Generate the edited deck and hand it to the browser as a download. */
  async #saveAs(): Promise<void> {
    if (!this.#presJson) return;
    const bytes = await this.savePresentation();
    const blob = new Blob([bytes as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
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

  /** The playable media frame under the slide point, resolved to the screen
   * box its overlay player docks at — null when the point hits nothing
   * playable. */
  #mediaPlaybackAt(point: { slide: number; x: number; y: number }): MediaPlayback | null {
    const pres = this.#pres;
    if (!pres) return null;
    const member = pres.slides[point.slide]?.members.find(
      (m) =>
        m.kind === "mediaFrame" &&
        m.playable &&
        point.x >= m.x &&
        point.x <= m.x + m.width &&
        point.y >= m.y &&
        point.y <= m.y + m.height,
    );
    if (!member || member.kind !== "mediaFrame" || !member.playable) return null;
    const surfaceRect = this.#canvasHost().getBoundingClientRect();
    const stageRect = this.#stage.getBoundingClientRect();
    if (stageRect.width === 0) return null;
    const scale = stageRect.width / pres.widthPx;
    return {
      media: member.media,
      src: member.playable.src,
      x: stageRect.left - surfaceRect.left + member.x * scale,
      y:
        stageRect.top -
        surfaceRect.top +
        (point.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX) * scale +
        member.y * scale,
      width: member.width * scale,
      height: member.height * scale,
    };
  }

  readonly #onStagePointerDown = (event: PointerEvent): void => {
    // A click advances the show (PowerPoint's rule) — no gestures inside.
    if (this.#presenting) {
      // A click on a playable frame toggles its playback instead of the
      // advance — the media rule outranks the slide rule.
      const point = event.button === 0 ? this.#stagePointOf(event) : null;
      const playback = point ? this.#mediaPlaybackAt(point) : null;
      if (playback) {
        const player = this.#mediaPlayer;
        if (player?.isOpenFor(playback.src)) player.toggle();
        else player?.show(playback);
        return;
      }
      this.#mediaPlayer?.hide();
      if (event.button === 0) this.#gotoShowSlide(this.#showingSlide + 1);
      return;
    }
    if (this.#shapeDrawer?.armed && this.#shapeDrawer.startFromPointer(event)) return;
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
      // Double-click on a playable media frame opens native playback (the
      // canvas poster keeps painting beneath the floating controls).
      const playback = this.#mediaPlaybackAt(point);
      if (playback) return this.#mediaPlayer?.show(playback);
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
    if (event.key === "Escape" && this.#shapeDrawer?.armed) {
      this.#shapeDrawer.disarm();
      return;
    }
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
      repaintSlide(
        this.#thumbApp.tree as unknown as IGroup,
        pres,
        slide,
        () => this.#thumbApp?.forceRender(),
        slide * pitch,
        this.#renderedThumbPres?.slides[slide],
      );
      this.#renderedThumbPres = pres;
      return;
    }
    this.#thumbApp?.destroy();
    this.#renderedThumbPres = null;
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
    this.#renderedThumbPres = pres;
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
