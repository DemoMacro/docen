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
  type SlideCommentOptions,
} from "@docen/pptx";
import { customElement, observable } from "@microsoft/fast-element";
import type { DataType } from "@office-open/core";
import type { ColorSchemeOptions, FontSchemeOptions } from "@office-open/core";
import type {
  RunFont,
  ShapeType,
  TextBodyOptions,
  TextFont,
  TextRunOptions,
  TextVertical,
} from "@office-open/core/drawing";
// Leafer ships animate() as a stub that only logs — the show's entrance
// tweens need the real plugin registered.
import "@leafer-in/animate";
import { App, type IGroup } from "leafer-ui";

import { buildContextualTab, renderRibbonFromSchema } from "../document/ribbon";
import {
  addSpellWord,
  englishWords,
  ignoreSpellWord,
  spellSuggestions,
} from "../document/spelling";
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
  makeSymbol,
  makeTable,
  makeWordArt,
  makeTextBox,
  makeFieldBox,
  makeLine,
  makeObject,
  makePenStroke,
  objectProgIdOf,
  makeMediaFrame,
  makeSmartArt,
  makeChart,
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
  THEME_PRESETS,
  variantSchemesOf,
  resetSlidePlaceholders,
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
import { TABLE_STYLE_IDS, tableDesignTab, tableLayoutTab } from "./ribbon";
import { ANIMATION_PRESETS, presentationRibbonTabs, TRANSITION_PRESETS } from "./ribbon";
import { changedSlides } from "./slide-diff";
import {
  paintSlideDeck,
  SLIDE_GAP_PX,
  repaintSlide,
  THUMB_GAP_PX,
  THUMB_WIDTH_PX,
} from "./slide-paint";
import { TableSelectionOverlay } from "./table-overlay";
import {
  tableGripAt,
  tableSelectionCells,
  tableCellAt,
  tableSelectionRects,
  tableCellNearAt,
  tableSelectionFor,
  type TableSelectionRange,
} from "./table-selection";
import {
  disposeTextAreaMirror,
  measureTextAreaContent,
  slideFillStyle,
  textBoxFillStyle,
} from "./text-overlay";
import { formatText, textOf, writeText, type TextEditSource } from "./text-session";
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

/** One misspelling in a PPTX text home. Targets keep live model references,
 *  so replacement writes through the same source used by in-place editing. */
interface PresentationSpellingIssue {
  slide: number;
  word: string;
  from: number;
  to: number;
  source: TextEditSource | { kind: "notes"; slide: SlideOptions };
}

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

/** Render the embedded object's browser-visible icon as a real PNG. The OLE
 *  payload stays source bytes; this preview is only the `p:pic` the file
 *  format requires. */
async function objectIconDataUrl(sourceName: string): Promise<string> {
  const extension = (sourceName.split(".").pop() ?? "FILE").toUpperCase().slice(0, 4);
  const safeExtension = extension.replace(/[<>&"']/g, "");
  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">' +
    '<path d="M24 8h54l26 26v86a8 8 0 0 1-8 8H24a8 8 0 0 1-8-8V16a8 8 0 0 1 8-8z" fill="#fff" stroke="#9aa5b1" stroke-width="4"/>' +
    '<path d="M78 8v26h26z" fill="#e8f0fe" stroke="#9aa5b1" stroke-width="4"/>' +
    `<text x="64" y="86" font-family="Segoe UI, sans-serif" font-size="28" font-weight="600" fill="#4472c4" text-anchor="middle">${safeExtension}</text>` +
    "</svg>";
  const image = new Image();
  image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await image.decode();
  const canvas = document.createElement("canvas");
  canvas.width = 128;
  canvas.height = 128;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("canvas unavailable");
  context.drawImage(image, 0, 0);
  return canvas.toDataURL("image/png");
}

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
  textVertical?: "vertical" | "vertical270";
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
  "save",
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
  "slide-sorter",
  "spell-check",
  "new-comment",
  "show-comments",
  "set-up-show",
  "transition",
  "effect-options",
  "apply-to-all",
  "animation-pane",
  "hyperlink",
  "from-beginning",
  "from-current",
  "animate",
  "add-animation",
  "slide-number",
  "date-time",
  "select",
  "draw-select",
  "draw-eraser",
  "draw-pen",
  "smartart",
  "chart",
  "symbol",
  "wordart",
  "online-picture",
  "object",
  "section",
  "header-footer",
  "layout",
  "reset",
  "themes",
  "variants",
  "toggle-table-look",
  "table-style",
  "cell-shading",
  "table-borders",
  "insert-above",
  "insert-below",
  "insert-left",
  "insert-right",
  "delete-table",
  "merge-cells",
  "split-cells",
  "cell-height",
  "cell-width",
  "distribute-rows",
  "distribute-columns",
  "table-align",
  "text-direction",
  "cell-margins",
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
  #tableOverlay: TableSelectionOverlay | null = null;
  #shapeDrawer: ShapeDrawer | null = null;
  #drawTool: "select" | "eraser" | "pen" | null = null;
  #mediaPlayer: MediaPlayer | null = null;
  /** The selected object: the slide child, or — with `member` set — the
   *  group-nested member it descended into (child indexes under the group). */
  #selection: { slide: number; child: number; member?: number[] } | null = null;
  /** Word's cross-cell selection for the selected table (grid row/col pairs;
   *  null collapses the selection to the cell next entered). */
  #tableSelection: TableSelectionRange | null = null;
  #tableTextSelection: HTMLDivElement[] = [];
  #contextTabIds = new Set<string>();
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
  /** The edited cell as it entered the session — live write-through mutates
   *  the cell in place, so this baseline is what the undo pair restores. */
  #tableEditBefore: TableCellOptions | null = null;
  /** The painted geometry the floating cell editor follows; re-projection
   *  refreshes it when auto row heights grow under live typing. */
  #tableEditLayout: {
    rect: { x: number; y: number; width: number; height: number };
    margins: { left: number; top: number; right: number; bottom: number };
    anchor: "top" | "center" | "bottom";
    baselineShift: number;
  } | null = null;
  /** Drawing gridlines visibility (a view state — not part of the deck). */
  #gridlines = false;
  /** The right task pane's current surface: PowerPoint's Selection and
   *  Animation panes share the same docking edge. */
  #rightPane: "animation" | "select" | null = null;
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
  /** PowerPoint's Slide Sorter: a modal browser over the existing thumbnail
   *  paints, with drag-and-drop deck reordering. */
  #slideSorter: HTMLDialogElement | null = null;
  #sorterDragIndex = -1;
  #sorterDragged = false;
  /** Review session state for PowerPoint's Spelling dialog. */
  #spellDialog: HTMLDialogElement | null = null;
  #spellIssues: PresentationSpellingIssue[] = [];
  #spellAt = 0;
  /** Session-level Set Up Show settings; the PPTX presentation properties
   *  part is not part of the public JSON yet. */
  #showSetup = {
    from: 1,
    to: Number.POSITIVE_INFINITY,
    loop: false,
    narration: true,
    animations: true,
    timings: true,
  };
  #showSetupDialog: HTMLDialogElement | null = null;
  #symbolDialog: HTMLDialogElement | null = null;
  #wordArtDialog: HTMLDialogElement | null = null;
  #commentsDialog: HTMLDialogElement | null = null;
  #sectionDialog: HTMLDialogElement | null = null;
  #headerFooterDialog: HTMLDialogElement | null = null;
  #onlinePictureDialog: HTMLDialogElement | null = null;
  #themesDialog: HTMLDialogElement | null = null;
  #variantsDialog: HTMLDialogElement | null = null;
  #layoutDialog: HTMLDialogElement | null = null;
  #objectDialog: HTMLDialogElement | null = null;
  #objectLink = false;
  #objectAutoUpdate = false;
  #objectShowAsIcon = true;
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
      .querySelector<HTMLInputElement>("#object-input")
      ?.addEventListener("change", this.#onObjectChange as EventListener);
    root
      .querySelector("docen-find-replace-dialog")
      ?.addEventListener("find-replace:action", this.#onFindReplace as EventListener);
    root
      .querySelector("docen-link-dialog")
      ?.addEventListener("link:ok", this.#onLinkOk as EventListener);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    document.addEventListener("selectionchange", this.#onDocumentSelectionChange);
    this.selectList?.addEventListener("click", this.#onSelectListClick);
    this.#area()?.addEventListener("wheel", this.#onShowWheel, { passive: false });
    this.#area()?.addEventListener("scroll", this.#onScroll);
    this.thumbStrip?.addEventListener("click", this.#onThumbClick);
    // Capture: Leafer's own app view can consume a pointer before bubbling;
    // the editor must get first refusal for an armed drag-to-draw tool.
    this.#stage.addEventListener("pointerdown", this.#onStagePointerDown, true);
    this.#stage.addEventListener("pointermove", this.#onStagePointerMove);
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
    this.#tableOverlay = new TableSelectionOverlay({
      scale: () => this.#zoom / 100,
      selectGrip: (kind, index) => this.#selectTableGrip(kind, index),
    });
    this.#canvasHost().append(this.#tableOverlay.el);
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
      ?.querySelector<HTMLInputElement>("#object-input")
      ?.removeEventListener("change", this.#onObjectChange as EventListener);
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
    this.#tableOverlay?.el.remove();
    this.#tableOverlay = null;
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
    // Transient dialogs hold closures over the previous deck's model — the
    // swap would leave them editing a detached copy.
    for (const dialog of this.shadowRoot?.querySelectorAll<HTMLDialogElement>("dialog[open]") ?? [])
      dialog.close();
    this.#symbolDialog = null;
    this.#onlinePictureDialog = null;
    this.#sectionDialog = null;
    this.#headerFooterDialog = null;
    this.#themesDialog = null;
    this.#variantsDialog = null;
    this.#layoutDialog = null;
    this.#objectDialog = null;
    this.#wordArtDialog = null;
    this.#spellDialog = null;
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
    this.#contextTabIds.clear();
    this.notesEditor?.setAttribute("placeholder", t("ppt.notes.placeholder", this));
    this.#applyRibbonGreying();
    this.#syncContextTabs();
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
    // QAT undo/redo follow the edit stack; Save hands the browser the current
    // package (the same Save As pipeline) whenever a deck is loaded.
    const qat = [
      { id: "save", icon: "save", disabled: !this.#presJson },
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
      "docen-ribbon-button[event], docen-ribbon-toggle-button[event], docen-ribbon-split-button[event], docen-ribbon-menu[event], docen-ribbon-input[event]",
    ) ?? []) {
      const event = el.getAttribute("event");
      if (!event || wired.has(event)) continue;
      el.setAttribute("disabled", "");
    }
  }

  /** PowerPoint's contextual table tabs: append/remove only when the selection
   *  enters/leaves a table, without rebuilding the whole ribbon. */
  #syncContextTabs(): void {
    const root = this.shadowRoot;
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    const scope = root.querySelector("docen-workspace") ?? this;
    const want = this.#tableMemberOf()
      ? new Map([
          ["ppt-table-design", tableDesignTab()],
          ["ppt-table-layout", tableLayoutTab()],
        ])
      : new Map();
    const present = this.#contextTabIds;
    const changed = want.size !== present.size || [...want.keys()].some((id) => !present.has(id));
    if (!changed) return;
    const active = tablist.getAttribute("activeid") ?? "";
    if (present.has(active) && !want.has(active)) tablist.setAttribute("activeid", "home");
    for (const id of present) {
      if (want.has(id)) continue;
      ribbon.querySelector(`docen-ribbon-panel[value="${id}"]`)?.remove();
      tablist.querySelector(`#${id}`)?.remove();
    }
    let firstNew: string | null = null;
    for (const [id, tab] of want) {
      if (present.has(id)) continue;
      const built = buildContextualTab(tab, scope);
      tablist.append(built.tab);
      ribbon.append(built.panel);
      firstNew = firstNew ?? id;
    }
    if (firstNew) tablist.setAttribute("activeid", firstNew);
    present.clear();
    for (const id of want.keys()) present.add(id);
    this.#applyRibbonGreying();
  }

  // ── Events ───────────────────────────────────────────────────────────────

  /** Title-bar menu items carry their action in `data-event`; the notes
   *  textarea's change (its blur commit) routes to the notes session. */
  readonly #onChange = (event: Event): void => {
    const target = event.target as HTMLElement;
    if (target === this.notesEditor) return this.#commitNotes();
    const animationControl = target.closest<HTMLElement>("[data-animation-setting]");
    if (animationControl) {
      this.#updateAnimationSetting(
        Number(animationControl.dataset.animationIndex),
        animationControl.dataset.animationSetting!,
        (target as HTMLInputElement | HTMLSelectElement).value,
      );
      return;
    }
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
    else if (name === "save") void this.#saveAs();
    else if (name === "save-as") void this.#saveAs();
    else if (name === "new-slide") this.#insertSlide();
    else if (name === "delete-slide") this.#deleteSlide();
    else if (name === "duplicate-slide") this.#duplicateSlide();
    else if (name === "text-box") this.#insertTextBox();
    else if (name === "insert-table") this.#insertTable();
    else if (name === "chart") this.#insertChart();
    else if (name === "shapes") this.#armShape(event.detail?.value);
    else if (name === "format-background") this.#setBackground(event.detail?.value);
    else if (name === "slide-size") this.#toggleSlideSize();
    else if (name === "insert-picture") this.#pickPicture();
    else if (name === "bring-front") this.#reorderSelected("front");
    else if (name === "send-back") this.#reorderSelected("back");
    else if (name === "gridlines") this.#toggleGridlines();
    else if (name === "find" || name === "replace") this.#findDialog()?.show();
    else if (name === "spell-check") this.#openSpellCheck();
    else if (name === "new-comment") this.#openCommentsDialog(true);
    else if (name === "show-comments") this.#toggleCommentsDialog();
    else if (name === "paste") void this.#pasteFromClipboard();
    else if (name === "notes") this.#toggleNotes();
    else if (name === "normal") {
      this.#slideSorter?.close();
      this.#enterNormalView();
    } else if (name === "slide-sorter") this.#toggleSlideSorter();
    else if (name === "transition" && event.detail?.value) this.#setTransition(event.detail.value);
    else if (name === "effect-options" && event.detail?.value) {
      this.#setTransitionSpeed(event.detail.value);
    } else if (name === "apply-to-all") this.#applyTransitionToAll();
    else if (name === "hyperlink") this.#openLinkDialog();
    else if (name === "from-beginning") this.#startShow("beginning");
    else if (name === "from-current") this.#startShow("current");
    else if (name === "set-up-show") this.#openShowSetup();
    else if (name === "animate") this.#applyAnimation(event.detail?.value, "replace");
    else if (name === "add-animation") this.#applyAnimation(event.detail?.value, "add");
    else if (name === "animation-pane") this.#toggleAnimationPane();
    else if (name === "slide-number" || name === "date-time") this.#insertFieldBox(name);
    else if (name === "smartart") this.#insertSmartArt();
    else if (name === "symbol") this.#openSymbolDialog();
    else if (name === "wordart") this.#openWordArtDialog();
    else if (name === "online-picture") this.#openOnlinePictureDialog();
    else if (name === "section") this.#openSectionDialog();
    else if (name === "header-footer") this.#openHeaderFooterDialog();
    else if (name === "themes") this.#openThemesDialog();
    else if (name === "variants") this.#openVariantsDialog();
    else if (name === "layout") this.#openLayoutPanel();
    else if (name === "reset") this.#resetSlideToLayout();
    else if (name === "video" || name === "audio") this.#pickMedia();
    else if (name === "object") this.#pickObject();
    else if (name === "select") this.#toggleSelectionPane();
    else if (name === "toggle-table-look") this.#toggleTableLook(event.detail?.value);
    else if (name === "table-style") this.#setTableStyle(event.detail?.value);
    else if (name === "cell-shading") this.#setTableCellShading(event.detail?.value);
    else if (name === "table-borders") this.#setTableCellBorders(event.detail?.value);
    else if (name === "insert-above") this.#insertTableRow("above");
    else if (name === "insert-below") this.#insertTableRow("below");
    else if (name === "insert-left") this.#insertTableColumn("left");
    else if (name === "insert-right") this.#insertTableColumn("right");
    else if (name === "delete-table") this.#deleteTablePart(event.detail?.value);
    else if (name === "merge-cells") this.#mergeSelectedTableCells();
    else if (name === "split-cells") this.#splitSelectedTableCell();
    else if (name === "cell-height") this.#setTableCellSize("height", event.detail?.value);
    else if (name === "cell-width") this.#setTableCellSize("width", event.detail?.value);
    else if (name === "distribute-rows") this.#distributeTableCells("rows");
    else if (name === "distribute-columns") this.#distributeTableCells("columns");
    else if (name === "table-align") this.#setTableCellAlignment(event.detail?.value);
    else if (name === "text-direction") this.#setTableCellTextDirection(event.detail?.value);
    else if (name === "cell-margins") this.#setTableCellMargins(event.detail?.value);
    else if (name === "draw-select") this.#armDrawTool("select");
    else if (name === "draw-eraser") this.#armDrawTool("eraser");
    else if (name === "draw-pen") this.#armDrawTool("pen");
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
    const index = this.#presenting
      ? this.#showingSlide
      : Math.max(0, Math.min(pres.slides.length - 1, Math.floor(center / pitch)));
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
    this.#syncSelectionOverlays();
    this.#renderSelectionPane();
    this.#syncTextFormatControls();
    this.#syncContextTabs();
  }

  /** Keep exactly one selection surface alive: tables get Word's cell grips
   *  and highlights; every other drawing keeps the transform frame. */
  #syncSelectionOverlays(): void {
    const table = this.#selection ? this.#tableMemberOf() : null;
    if (table) {
      this.#tableOverlay?.show(table.member, this.#tableStripY(table.slide), this.#tableSelection);
      this.#overlay?.hide();
      return;
    }
    this.#tableOverlay?.hide();
    this.#tableSelection = null;
    const hit = this.#selection ? this.#selectedBox() : null;
    if (hit) this.#overlay?.show(hit.box, hit.rotation);
    else this.#overlay?.hide();
  }

  /** The slide's top offset in the strip (slide-local y → overlay y). */
  #tableStripY(slide: number): number {
    const pres = this.#pres!;
    return slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
  }

  /** Set the grid pair and repaint only the table's cell highlights. */
  #setTableSelection(range: TableSelectionRange | null): void {
    this.#tableSelection = range;
    this.#tableOverlay?.setSelection(range);
    this.#syncTableCellSizeControls();
    if (range) this.#clearTableCellTextSelection();
    else this.#syncTableCellTextSelection();
  }

  /** A grip click commits the current cell first, then widens Word-style:
   *  top strips select columns, left strips rows, the corner selects all. */
  #selectTableGrip(kind: "row" | "col" | "table", index = 0): void {
    if (this.#textEditor) this.#exitTextEditing(true);
    const table = this.#tableMemberOf();
    const range = table ? tableSelectionFor(table.member, kind, index) : null;
    if (range) this.#setTableSelection(range);
  }

  /** DOCX's hover pass: one table grip resolves and paints at a time, so the
   *  table no longer carries a permanent fence of invisible hit boxes. */
  readonly #onStagePointerMove = (event: PointerEvent): void => {
    // A live cell edit keeps the edge/corner grips, but its body hover square
    // would read as an unrelated frame over the caret.
    if (!this.#selection) return;
    const table = this.#tableMemberOf();
    const point = this.#stagePointOf(event);
    if (!table || !point || point.slide !== table.slide) {
      this.#tableOverlay?.hover(Number.NaN, Number.NaN);
      return;
    }
    this.#tableOverlay?.hover(
      point.x - table.member.x,
      point.y - table.member.y,
      Boolean(this.#textEditor),
    );
  };

  readonly #onDocumentSelectionChange = (): void => {
    const editor = this.#textEditor;
    if (editor?.classList.contains("table-cell-editor")) this.#syncTableCellTextSelection(editor);
  };

  /** Word's cell-selection Delete: clear content from every selected origin,
   *  while the grid and its formatting survive. */
  #clearSelectedTableCells(): void {
    const table = this.#tableMemberOf();
    const range = this.#tableSelection;
    if (!table || !range) return;
    const origins = tableGridOf(table.table).origins;
    const selected = tableSelectionCells(table.member, range)
      .map(
        ({ row, col }) => origins.find((origin) => origin.row === row && origin.col === col)?.cell,
      )
      .filter((cell): cell is TableCellOptions => Boolean(cell));
    const before = selected.map((cell) => ({
      cell,
      text: cell.text,
      children: structuredClone(cell.children),
    }));
    if (before.every(({ cell }) => textOf({ kind: "cell", cell }) === "")) return;
    for (const { cell } of before) writeText({ kind: "cell", cell }, "");
    this.#pushEdit({
      undo: () => {
        for (const snapshot of before) {
          delete snapshot.cell.text;
          delete snapshot.cell.children;
          if (snapshot.text !== undefined) snapshot.cell.text = snapshot.text;
          if (snapshot.children !== undefined) snapshot.cell.children = snapshot.children;
        }
        this.#reproject(table.slide);
      },
      redo: () => {
        for (const { cell } of before) writeText({ kind: "cell", cell }, "");
        this.#reproject(table.slide);
      },
    });
    this.#reproject(table.slide);
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
      : slideFillStyle(slide.background, pres.widthPx, pres.heightPx, {
          x: box.x,
          y: box.y,
          scale,
        });
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
    this.#syncSelectionOverlays();
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
    const rect = tableCellAt(member, x, y);
    return rect ? { ...rect, cell: rect.cell as TableCellView } : null;
  }

  /** Float a textarea over one cell of the selected table (PowerPoint's
   *  in-place cell edit): insets, face, paragraph alignment and vertical
   *  anchor come from the painted cell; the text writes back on exit, and a
   *  click on another cell moves the session there. */
  #enterTableCellEditing(
    x: number,
    y: number,
    range?: TableSelectionRange,
    caret?: { clientX: number; clientY: number },
    selectAll = false,
  ): void {
    const found = this.#tableMemberOf();
    if (!found) return;
    const rect = this.#cellRectAt(found.member, x, y);
    if (!rect) return;
    this.#setTableSelection(range ?? null);
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
    editor.className = "shape-text-editor table-cell-editor";
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
            color: "transparent",
            caretColor: run.color ? `#${run.color}` : TEXT_INK,
            ...(run.bold ? { fontWeight: "bold" } : {}),
            ...(run.italic ? { fontStyle: "italic" } : {}),
          }
        : {
            color: "transparent",
            fontSize: `${((firstCellRunSizeOf(source) * 4) / 3) * scale}px`,
            caretColor: TEXT_INK,
          }),
      ...(cell.textVertical
        ? {
            writingMode: cell.textVertical === "vertical270" ? "sideways-lr" : "vertical-rl",
          }
        : {}),
    });
    editor.addEventListener("keydown", (event) => {
      // Escape leaves the edit (committing); Tab walks cells like DOCX and
      // every other key stays in the textarea's native text session.
      if (this.#tableSelection) {
        if (event.key === "Escape" || event.key === "Enter") {
          event.preventDefault();
          event.stopPropagation();
          this.#setTableSelection(null);
          return;
        }
        if (["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key)) {
          event.preventDefault();
          event.stopPropagation();
          this.#setTableSelection(null);
          const edge =
            event.key === "ArrowLeft" || event.key === "ArrowUp" ? 0 : editor.value.length;
          editor.setSelectionRange(edge, edge);
          return;
        }
        if (
          (event.key === "Delete" || event.key === "Backspace") &&
          !event.ctrlKey &&
          !event.metaKey &&
          !event.altKey
        ) {
          event.preventDefault();
          event.stopPropagation();
          this.#clearSelectedTableCells();
          editor.value = "";
          this.#layoutTableCellEditor(editor);
          return;
        }
        // Typing collapses the block to the caret (Word's rule) — the
        // leftover highlights would otherwise ghost over the edited cell.
        if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
          this.#clearSelectedTableCells();
          editor.value = "";
          this.#layoutTableCellEditor(editor);
          this.#setTableSelection(null);
        }
      }
      if (event.key === "Tab") {
        event.stopPropagation();
        event.preventDefault();
        this.#moveTableCellEditor(event.shiftKey ? -1 : 1);
        return;
      }
      if (event.key === "Escape") {
        event.stopPropagation();
        this.#exitTextEditing(true);
      }
      event.stopPropagation();
    });
    let composing = false;
    let lastWritten = editor.value;
    const writeThrough = (): void => {
      if (composing || this.#textEditRaf) return;
      this.#textEditRaf = requestAnimationFrame(() => {
        this.#textEditRaf = 0;
        if (this.#textEditor !== editor || editor.value === lastWritten) return;
        lastWritten = editor.value;
        writeText({ kind: "cell", cell: source }, editor.value);
        this.#reproject(found.slide);
      });
    };
    editor.addEventListener("input", () => {
      this.#layoutTableCellEditor(editor);
      writeThrough();
      this.#syncTableCellTextSelection(editor);
    });
    editor.addEventListener("compositionstart", () => {
      composing = true;
    });
    editor.addEventListener("compositionend", () => {
      composing = false;
      this.#layoutTableCellEditor(editor);
      writeThrough();
    });
    this.#stage.append(editor);
    this.#clearTableCellTextSelection();
    this.#tableEditLayout = {
      rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
      margins: m,
      anchor,
      baselineShift,
    };
    this.#layoutTableCellEditor(editor);
    editor.focus();
    if (selectAll) editor.select();
    else if (caret) this.#placeTableCellCaret(editor, caret.clientX, caret.clientY);
    else editor.setSelectionRange(editor.value.length, editor.value.length);
    this.#textEditor = editor;
    this.#tableEdit = { row: cell.row, col: cell.col };
    this.#tableEditBefore = structuredClone(source);
    this.#syncTableCellSizeControls();
    this.#syncTableCellTextSelection();
    this.#restoreOverlay();
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
    if (this.#textEditRaf) {
      cancelAnimationFrame(this.#textEditRaf);
      this.#textEditRaf = 0;
    }
    editor?.remove();
    disposeTextAreaMirror(editor);
    const sel = this.#selection;
    const before = this.#tableEditBefore;
    this.#tableEditBefore = null;
    this.#tableEditLayout = null;
    const cell = editor && write ? this.#editedCell(at) : null;
    if (cell && before) {
      const source: TextEditSource = { kind: "cell", cell };
      if (editor!.value !== textOf(source)) writeText(source, editor!.value);
      const after = structuredClone(cell);
      if (JSON.stringify(before) !== JSON.stringify(after)) {
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
    this.#clearTableCellTextSelection();
  }

  /** Re-fit the floating editor after a projection refresh: the painted row
   *  can grow while the textarea still carries the previous frame. */
  #layoutTableCellEditor(editor: HTMLTextAreaElement): void {
    const layout = this.#tableEditLayout;
    if (!layout) return;
    const scale = this.#zoom / 100;
    const { rect, margins: m, anchor, baselineShift } = layout;
    editor.style.left = `${rect.x * scale}px`;
    editor.style.top = `${(rect.y + this.#tableStripY(this.#selection!.slide)) * scale}px`;
    editor.style.width = `${rect.width * scale}px`;
    editor.style.padding = `${m.top * scale}px ${m.right * scale}px ${m.bottom * scale}px ${m.left * scale}px`;
    editor.style.height = "auto";
    editor.style.paddingTop = `${m.top * scale + baselineShift}px`;
    const boxHeight = rect.height * scale;
    const paddingTB = (m.top + m.bottom) * scale;
    const contentHeight = measureTextAreaContent(editor);
    if (anchor === "top") {
      editor.style.height = `${Math.max(boxHeight, contentHeight + paddingTB)}px`;
    } else {
      const slack = boxHeight - paddingTB - contentHeight;
      if (slack >= 0) {
        editor.style.height = `${boxHeight}px`;
        editor.style.paddingTop = `${m.top * scale + baselineShift + (anchor === "center" ? slack / 2 : slack)}px`;
      } else {
        editor.style.height = `${contentHeight + paddingTB}px`;
      }
    }
    editor.scrollTop = 0;
    this.#syncTableCellTextSelection();
  }

  /** A click on another cell of the table under edit: commit the current
   *  cell and float the editor over the new one (PowerPoint's cell hop). */
  #moveTableCellEditing(
    x: number,
    y: number,
    range?: TableSelectionRange,
    options?: { caret?: { clientX: number; clientY: number }; selectAll?: boolean },
  ): void {
    const found = this.#tableMemberOf();
    const rect = found ? this.#cellRectAt(found.member, x, y) : null;
    if (!rect) return this.#exitTextEditing(true);
    if (this.#tableEdit?.row === rect.cell.row && this.#tableEdit.col === rect.cell.col) return;
    this.#exitTextEditing(true);
    this.#enterTableCellEditing(x, y, range, options?.caret, options?.selectAll);
  }

  /** DOCX's Tab/Shift+Tab: move to the next/previous origin cell and select
   *  that cell's content so typing replaces it. */
  #moveTableCellEditor(delta: number): void {
    const found = this.#tableMemberOf();
    const at = this.#tableEdit;
    if (!found || !at) return;
    const origins = tableGridOf(found.table).origins;
    const current = origins.findIndex((origin) => origin.row === at.row && origin.col === at.col);
    const target = origins[current + delta];
    if (!target) return this.#exitTextEditing(true);
    const point = this.#tableCellCenter(found.member, target);
    this.#moveTableCellEditing(point.x, point.y, undefined, { selectAll: true });
  }

  /** A grid cell's slide-local center — enough for the origin-resolving cell
   *  hit test, including merged cells. */
  #tableCellCenter(
    member: TableMemberView,
    cell: { row: number; col: number; spanW: number; spanH: number },
  ): { x: number; y: number } {
    const left = member.table.columnWidthsPx
      .slice(0, cell.col)
      .reduce((sum, width) => sum + width, 0);
    const width = member.table.columnWidthsPx
      .slice(cell.col, cell.col + cell.spanW)
      .reduce((sum, width) => sum + width, 0);
    const top = member.table.rows.slice(0, cell.row).reduce((sum, row) => sum + row.heightPx, 0);
    const height = member.table.rows
      .slice(cell.row, cell.row + cell.spanH)
      .reduce((sum, row) => sum + row.heightPx, 0);
    return { x: member.x + left + width / 2, y: member.y + top + height / 2 };
  }

  /** Put the native textarea caret where the pointer met the canvas. The
   *  browser resolves text offsets for the just-focused control; a fallback
   *  leaves the caret at the text end rather than DOCX's odd select-all. */
  #placeTableCellCaret(editor: HTMLTextAreaElement, clientX: number, clientY: number): void {
    const position = document.caretPositionFromPoint?.(clientX, clientY);
    if (position?.offsetNode === editor && typeof position.offset === "number") {
      editor.setSelectionRange(position.offset, position.offset);
      return;
    }
    const range = document.caretRangeFromPoint?.(clientX, clientY);
    if (range?.startContainer === editor && typeof range.startOffset === "number") {
      editor.setSelectionRange(range.startOffset, range.startOffset);
      return;
    }
    const offset = this.#textOffsetAtPoint(editor, clientX, clientY);
    editor.setSelectionRange(offset, offset);
  }

  /** DOCX paints text selection in the canvas layer; the cell bridge's text is
   *  transparent, so mirror the native selection into overlay bands instead. */
  #syncTableCellTextSelection(editor = this.#textEditor): void {
    if (
      !editor ||
      editor !== this.#textEditor ||
      !this.#tableEdit ||
      this.#tableSelection ||
      editor.selectionStart === editor.selectionEnd
    ) {
      return this.#clearTableCellTextSelection();
    }
    const style = getComputedStyle(editor);
    const box = editor.getBoundingClientRect();
    const mirror = document.createElement("div");
    mirror.setAttribute("aria-hidden", "true");
    Object.assign(mirror.style, {
      position: "fixed",
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      boxSizing: style.boxSizing,
      padding: style.padding,
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      lineHeight: style.lineHeight,
      letterSpacing: style.letterSpacing,
      textAlign: style.textAlign,
      whiteSpace: "pre-wrap",
      overflowWrap: style.overflowWrap,
      wordBreak: style.wordBreak,
      tabSize: style.tabSize,
      direction: style.direction,
      visibility: "hidden",
      pointerEvents: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    mirror.textContent = editor.value;
    document.body.append(mirror);
    const text = mirror.firstChild;
    const rects: DOMRect[] = [];
    if (text && editor.selectionEnd > editor.selectionStart) {
      const range = document.createRange();
      range.setStart(text, editor.selectionStart);
      range.setEnd(text, editor.selectionEnd);
      rects.push(...range.getClientRects());
    }
    mirror.remove();
    if (rects.length === 0) return this.#clearTableCellTextSelection();
    const hostRect = this.#canvasHost().getBoundingClientRect();
    while (this.#tableTextSelection.length < rects.length) {
      const band = document.createElement("div");
      band.className = "table-text-selection";
      this.#canvasHost().append(band);
      this.#tableTextSelection.push(band);
    }
    this.#tableTextSelection.forEach((band, index) => {
      const rect = rects[index];
      if (!rect || rect.width === 0 || rect.height === 0) {
        band.style.display = "none";
        return;
      }
      Object.assign(band.style, {
        display: "block",
        left: `${rect.left - hostRect.left}px`,
        top: `${rect.top - hostRect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      });
    });
  }

  #clearTableCellTextSelection(): void {
    for (const band of this.#tableTextSelection) band.remove();
    this.#tableTextSelection = [];
  }

  /** The text offset at a viewport point when the browser can't resolve a
   *  caret inside this shadow-DOM textarea: the control's own font/box is
   *  cloned off-screen and per-character Range boxes pick the nearest side. */
  #textOffsetAtPoint(editor: HTMLTextAreaElement, clientX: number, clientY: number): number {
    const value = editor.value;
    if (!value) return 0;
    const style = getComputedStyle(editor);
    const box = editor.getBoundingClientRect();
    const mirror = document.createElement("div");
    Object.assign(mirror.style, {
      position: "fixed",
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      boxSizing: style.boxSizing,
      padding: style.padding,
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      lineHeight: style.lineHeight,
      letterSpacing: style.letterSpacing,
      textAlign: style.textAlign,
      whiteSpace: "pre-wrap",
      overflowWrap: style.overflowWrap,
      wordBreak: style.wordBreak,
      tabSize: style.tabSize,
      direction: style.direction,
      visibility: "hidden",
      pointerEvents: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    mirror.textContent = value;
    document.body.append(mirror);
    const text = mirror.firstChild;
    let best = 0;
    let bestDistance = Number.POSITIVE_INFINITY;
    if (text) {
      for (let index = 0; index < value.length; index += 1) {
        const range = document.createRange();
        range.setStart(text, index);
        range.setEnd(text, index + 1);
        const rect = range.getBoundingClientRect();
        if (!rect.width && !rect.height) continue;
        const verticalDistance =
          clientY < rect.top
            ? rect.top - clientY
            : clientY > rect.bottom
              ? clientY - rect.bottom
              : 0;
        const characterMiddle = rect.left + rect.width / 2;
        const offset = clientX < characterMiddle ? index : index + 1;
        const horizontalDistance = Math.abs(clientX - characterMiddle);
        const distance = verticalDistance * 10000 + horizontalDistance;
        if (distance < bestDistance) {
          best = offset;
          bestDistance = distance;
        }
      }
    }
    mirror.remove();
    return best;
  }

  /** Cross-cell drag: the anchor cell keeps the edit session while every new
   *  grid slot under the pointer extends a whole-cell selection; returning to
   *  the anchor collapses back to the native text selection. */
  #startTableCellDrag(
    event: PointerEvent,
    found: { slide: number; table: TableOptions; member: TableMemberView },
    onTap?: () => void,
  ): void {
    if (!found) return;
    // The cell-selection drag owns the gesture; without this the textarea's
    // native selection drag suppresses the pointer stream and the highlight
    // freezes at its first cell.
    event.preventDefault();
    const origin = this.#tableEdit;
    const anchor = origin ?? this.#tableSelection?.anchor ?? null;
    const drag = { startX: event.clientX, startY: event.clientY, moved: false };
    const onMove = (move: PointerEvent): void => {
      if (!drag.moved && Math.hypot(move.clientX - drag.startX, move.clientY - drag.startY) < 3)
        return;
      drag.moved = true;
      const point = this.#stagePointOf(move);
      if (!point || point.slide !== found.slide) return;
      const head = tableCellNearAt(found.member, point.x, point.y);
      if (!head) return;
      if (origin && head.cell.row === origin.row && head.cell.col === origin.col) {
        if (this.#tableSelection) this.#setTableSelection(null);
        return;
      }
      if (!anchor) return;
      this.#setTableSelection({
        anchor,
        head: { row: head.cell.row, col: head.cell.col },
      });
    };
    const onUp = (): void => {
      document.removeEventListener("pointermove", onMove, { capture: true });
      document.removeEventListener("pointerup", onUp, { capture: true });
      document.removeEventListener("pointercancel", onUp, { capture: true });
      this.#syncTableCellTextSelection();
      if (!drag.moved) onTap?.();
    };
    document.addEventListener("pointermove", onMove, { capture: true });
    document.addEventListener("pointerup", onUp, { capture: true });
    document.addEventListener("pointercancel", onUp, { capture: true });
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
    // The paint restarted under the same selection — the table highlights or
    // transform frame snaps to the current geometry (and rejects itself if
    // the box vanished).
    if (this.#selectedBox()) this.#syncSelectionOverlays();
    else this.#select(null);
    const tableEditor = this.#textEditor;
    const at = this.#tableEdit;
    if (tableEditor && at && this.#tableEditLayout) {
      const member = this.#tableMemberOf()?.member;
      const cell = member?.table.rows
        .flatMap((row) => row.cells)
        .find((cell) => cell.row === at.row && cell.col === at.col);
      const rect =
        member && cell ? tableSelectionRects(member, { anchor: cell, head: cell })[0] : null;
      if (rect) {
        this.#tableEditLayout.rect = {
          x: rect.x,
          y: rect.y,
          width: rect.width,
          height: rect.height,
        };
        this.#layoutTableCellEditor(tableEditor);
      }
    }
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

  /** The editable array behind the current selection: the slide for a top-level
   *  child, or the nearest group container for a member path. */
  #selectedChildren(): { children: SlideChild[]; index: number } | null {
    const sel = this.#selection;
    const slide = this.#presJson?.slides?.[sel?.slide ?? -1];
    const children = slide?.children;
    if (!sel || !children) return null;
    if (!sel.member) return { children, index: sel.child };
    let node: SlideChild | undefined = children[sel.child];
    for (const index of sel.member.slice(0, -1)) {
      if (!node || !("group" in node)) return null;
      node = node.group.children?.[index];
    }
    if (!node || !("group" in node)) return null;
    const members = node.group.children;
    const index = sel.member[sel.member.length - 1] ?? -1;
    return members && index >= 0 && index < members.length ? { children: members, index } : null;
  }

  /** Remove the selected object; undo re-splices the same child back. */
  #deleteSelected(): void {
    const sel = this.#selection;
    const selected = this.#selectedChildren();
    if (!sel || !selected) return;
    const { children, index } = selected;
    const [child] = children.splice(index, 1);
    if (!child) return;
    this.#select(null);
    this.#pushEdit({
      undo: () => {
        children.splice(index, 0, child);
        this.#reproject(sel.slide);
      },
      redo: () => {
        children.splice(index, 1);
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
    this.#renderSlideSorter();
  }

  #redo(): void {
    if (!this.#history.canRedo) return;
    this.#history.redo();
    this.#exitTextEditing(true);
    this.#syncQat();
    this.#renderSlideSorter();
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
    const selected = this.#selectedChildren();
    if (!sel || !selected || selected.children.length < 2) return;
    const { children, index: from } = selected;
    const moved = to === "front" ? children.length - 1 : 0;
    reorderChild(children, from, moved);
    this.#select(
      sel.member ? { ...sel, member: [...sel.member] } : { slide: sel.slide, child: moved },
    );
    this.#pushEdit({
      undo: () => {
        reorderChild(children, moved, from);
        this.#select(
          sel.member ? { ...sel, member: [...sel.member] } : { slide: sel.slide, child: from },
        );
        this.#reproject(sel.slide);
      },
      redo: () => {
        reorderChild(children, from, moved);
        this.#select(
          sel.member ? { ...sel, member: [...sel.member] } : { slide: sel.slide, child: moved },
        );
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

  #insertChart(): void {
    const pres = this.#pres;
    if (!pres) return;
    this.#insertChild(makeChart(pres.widthPx, pres.heightPx));
  }

  /** A compact Unicode palette plus a direct character field. The insertion
   *  is an ordinary shape, so the symbol stays editable and undoable. */
  #openSymbolDialog(): void {
    if (!this.#pres || this.#symbolDialog?.open) return;
    const symbols = [
      "×",
      "÷",
      "±",
      "≤",
      "≥",
      "≠",
      "∞",
      "√",
      "∑",
      "π",
      "©",
      "®",
      "™",
      "°",
      "→",
      "←",
      "↑",
      "↓",
      "★",
      "☆",
      "♥",
      "✓",
      "✗",
      "•",
    ];
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.symbol.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><div class="symbol-grid">${symbols.map((symbol) => `<button data-symbol="${escapeHtml(symbol)}">${escapeHtml(symbol)}</button>`).join("")}</div><label class="dialog-field"><span>${escapeHtml(t("ppt.symbol.custom", this))}</span><input id="symbol-value" maxlength="8" value="©"></label></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button><button data-symbol-insert>${escapeHtml(t("ppt.dialog.insert", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      const symbol = target.closest<HTMLElement>("[data-symbol]")?.dataset.symbol;
      if (symbol) {
        this.#insertChild(makeSymbol(this.#pres!.widthPx, this.#pres!.heightPx, symbol));
        dialog.close();
      } else if (target.closest("[data-symbol-insert]")) {
        const value = dialog.querySelector<HTMLInputElement>("#symbol-value")?.value;
        if (value) {
          this.#insertChild(makeSymbol(this.#pres!.widthPx, this.#pres!.heightPx, value));
          dialog.close();
        }
      }
    });
    this.#showModalDialog(dialog, () => (this.#symbolDialog = null));
    this.#symbolDialog = dialog;
  }

  /** Fresh WordArt: a transparent, centered display-text box selected for
   *  immediate in-place editing. */
  #openWordArtDialog(): void {
    if (!this.#pres || this.#wordArtDialog?.open) return;
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.wordart.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><label class="dialog-field"><span>${escapeHtml(t("ppt.wordart.text", this))}</span><input id="wordart-value" value="${escapeHtml(t("ppt.wordart.default", this))}"></label></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button><button data-wordart-insert>${escapeHtml(t("ppt.dialog.insert", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (target.closest("[data-wordart-insert]")) {
        const value = dialog.querySelector<HTMLInputElement>("#wordart-value")?.value.trim();
        if (value) {
          this.#insertChild(makeWordArt(this.#pres!.widthPx, this.#pres!.heightPx, value));
          dialog.close();
        }
      }
    });
    this.#showModalDialog(dialog, () => (this.#wordArtDialog = null));
    this.#wordArtDialog = dialog;
  }

  /** Online pictures persist as real image bytes in the PPTX: the URL is
   *  fetched, typed by MIME, converted to a data URL, measured and inserted. */
  #openOnlinePictureDialog(): void {
    if (!this.#pres || this.#onlinePictureDialog?.open) return;
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.online-picture.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><label class="dialog-field"><span>${escapeHtml(t("ppt.online-picture.url", this))}</span><input id="online-picture-url" type="url" placeholder="https://"></label></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button><button data-picture-insert>${escapeHtml(t("ppt.dialog.insert", this))}</button></div>
    `;
    dialog.addEventListener("click", async (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (!target.closest("[data-picture-insert]")) return;
      const url = dialog.querySelector<HTMLInputElement>("#online-picture-url")?.value.trim();
      if (!url) return;
      target.setAttribute("disabled", "");
      try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const blob = await response.blob();
        const type = PICTURE_TYPES.get(blob.type);
        if (!type) throw new Error(blob.type || "unsupported image");
        const data = await readFileAsDataURL(new File([blob], "online-image", { type: blob.type }));
        const image = new Image();
        image.src = data;
        await image.decode();
        this.#insertChild(
          makePicture(
            this.#pres!.widthPx,
            this.#pres!.heightPx,
            image.naturalWidth,
            image.naturalHeight,
            data,
            type,
          ),
        );
        dialog.close();
      } catch (error) {
        window.alert(
          `${t("ppt.online-picture.error", this)}: ${error instanceof Error ? error.message : String(error)}`,
        );
      } finally {
        target.removeAttribute("disabled");
      }
    });
    this.#showModalDialog(dialog, () => (this.#onlinePictureDialog = null));
    this.#onlinePictureDialog = dialog;
  }

  /** Slide sections are stored on each member slide; the compiler groups the
   *  shared name into presentation.xml's sectionLst on save. */
  #openSectionDialog(): void {
    if (!this.#pres || this.#sectionDialog?.open) return;
    const slide = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!slide) return;
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.section.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body">
        <label class="dialog-field"><span>${escapeHtml(t("ppt.section.name", this))}</span><input id="section-name" value="${escapeHtml(slide.section ?? "")}" placeholder="${escapeHtml(t("ppt.section.placeholder", this))}"></label>
      </div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button><button data-section-save>${escapeHtml(t("ppt.dialog.insert", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (target.closest("[data-section-save]")) {
        const value = dialog.querySelector<HTMLInputElement>("#section-name")?.value.trim();
        this.#setSlideSection(value || undefined);
        dialog.close();
      }
    });
    this.#showModalDialog(dialog, () => (this.#sectionDialog = null));
    this.#sectionDialog = dialog;
    dialog.querySelector<HTMLInputElement>("#section-name")?.select();
  }

  #setSlideSection(section?: string): void {
    const slide = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!slide || slide.section === section) return;
    const before = slide.section;
    const apply = (value?: string): void => {
      if (value) slide.section = value;
      else delete slide.section;
    };
    apply(section);
    this.#pushEdit({
      undo: () => apply(before),
      redo: () => apply(section),
    });
  }

  /** The accent swatch strip every theme card shows. */
  #themeSwatches(scheme: ColorSchemeOptions): string {
    const accents = [
      scheme.accent1,
      scheme.accent2,
      scheme.accent3,
      scheme.accent4,
      scheme.accent5,
      scheme.accent6,
    ];
    return accents
      .map(
        (color) =>
          `<i style="background:#${typeof color === "string" ? color : (color?.lastClr ?? "FFFFFF")}"></i>`,
      )
      .join("");
  }

  /** The theme galleries write the master's real theme (colors, and the
   *  presets' font pair too) — the projection reads that same scheme, so the
   *  canvas repaints scheme-driven fills and table styles deck-wide. */
  #openThemesDialog(): void {
    if (!this.#pres || this.#themesDialog?.open) return;
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.themes.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><div class="theme-grid">
        ${THEME_PRESETS.map(
          (preset) => `
          <button class="theme-card" data-theme="${preset.id}">
            <span class="theme-swatches">${this.#themeSwatches(preset.colorScheme)}</span>
            <span class="theme-name">${escapeHtml(preset.colorScheme.name ?? preset.id)}</span>
            <span class="theme-fonts">${escapeHtml(`${preset.fontScheme?.majorFont?.latin?.typeface ?? ""} / ${preset.fontScheme?.minorFont?.latin?.typeface ?? ""}`)}</span>
          </button>`,
        ).join("")}
      </div></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      const id = target.closest("[data-theme]")?.getAttribute("data-theme");
      const preset = THEME_PRESETS.find((candidate) => candidate.id === id);
      if (preset) {
        this.#applyThemeColors(preset.colorScheme, preset.fontScheme);
        dialog.close();
      }
    });
    this.#showModalDialog(dialog, () => (this.#themesDialog = null));
    this.#themesDialog = dialog;
  }

  /** Variants recolor the deck's current scheme: the accents rotate so each
   *  leads in turn, exactly four choices like PowerPoint's variant strip. */
  #openVariantsDialog(): void {
    if (!this.#pres || this.#variantsDialog?.open) return;
    const current = this.#presJson?.masters?.[0]?.theme?.colorScheme;
    if (!current) return;
    const variants = variantSchemesOf(current);
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.variants.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><div class="theme-grid">
        ${variants
          .map(
            (variant, index) => `
          <button class="theme-card" data-variant="${index}">
            <span class="theme-swatches">${this.#themeSwatches(variant)}</span>
            <span class="theme-name">${escapeHtml(variant.name ?? `#${index + 1}`)}</span>
          </button>`,
          )
          .join("")}
      </div></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      const index = Number(target.closest("[data-variant]")?.getAttribute("data-variant"));
      const variant = variants[index];
      if (variant) {
        this.#applyThemeColors(variant);
        dialog.close();
      }
    });
    this.#showModalDialog(dialog, () => (this.#variantsDialog = null));
    this.#variantsDialog = dialog;
  }

  /** Write one theme into the first master (creating the master entry when a
   *  fresh deck has none — the compiler's implicit default master is the same
   *  slot), with undo/redo swapping the whole previous theme back. */
  #applyThemeColors(colorScheme: ColorSchemeOptions, fontScheme?: FontSchemeOptions): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    if (!presJson.masters) presJson.masters = [{}];
    const master = presJson.masters[0]!;
    const before = master.theme ? structuredClone(master.theme) : undefined;
    const after = {
      ...before,
      colorScheme: structuredClone(colorScheme),
      ...(fontScheme ? { fontScheme: structuredClone(fontScheme) } : {}),
    };
    master.theme = after;
    this.#pushEdit({
      undo: () => {
        if (before) master.theme = structuredClone(before);
        else delete master.theme;
        this.#reproject();
      },
      redo: () => {
        master.theme = structuredClone(after);
        this.#reproject();
      },
    });
    this.#reproject();
  }

  /** PowerPoint's Layout gallery: switching the slide's layout also re-seats
   *  the placeholders it inherits (the same walk Reset runs). The gallery
   *  lists the master's real layouts — the compiler keys slide→layout by
   *  type/name, so offering absent layouts would silently keep the old part. */
  #openLayoutPanel(): void {
    if (!this.#pres || this.#layoutDialog?.open) return;
    const presJson = this.#presJson;
    const slideIndex = this.#activeSlideIndex();
    const slide = presJson?.slides?.[slideIndex];
    if (!presJson || !slide) return;
    const master =
      presJson.masters?.find((m) => m.name === (slide.master ?? m.name)) ?? presJson.masters?.[0];
    const fromMaster = (master?.layouts ?? [])
      .map((candidate) => ({ key: candidate.type ?? candidate.name ?? "", name: candidate.name }))
      .filter((candidate) => candidate.key !== "");
    const LAYOUTS =
      fromMaster.length > 0
        ? fromMaster
        : [
            "title",
            "titleOnly",
            "blank",
            "text",
            "twoColumnText",
            "object",
            "sectionHeader",
            "twoObjects",
            "objectAndText",
            "pictureText",
            "clipArtAndText",
            "twoTextAndTwoObjects",
            "verticalText",
            "verticalTitleAndText",
            "chart",
            "table",
          ].map((key) => ({ key, name: undefined as string | undefined }));
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.layout.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body"><div class="layout-grid">
        ${LAYOUTS.map(
          (layout) => `
          <button class="layout-card${slide.layout === layout.key ? " current" : ""}" data-layout="${escapeHtml(layout.key)}">
            <span class="layout-name">${escapeHtml(layout.name ?? t(`ppt.layout.type.${layout.key}`, this))}</span>
          </button>`,
        ).join("")}
      </div></div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      const type = target.closest("[data-layout]")?.getAttribute("data-layout") ?? undefined;
      if (!type || type === slide.layout) return;
      const before = slide.layout;
      const apply = (value: string | undefined): void => {
        if (value) slide.layout = value;
        else delete slide.layout;
      };
      const reseat = (): void => {
        const target = master?.layouts?.find(
          (candidate) => candidate.type === slide.layout || candidate.name === slide.layout,
        );
        for (const item of resetSlidePlaceholders(slide, target, master)) {
          Object.assign(item.child.shape, item.after);
        }
      };
      apply(type);
      reseat();
      this.#pushEdit({
        undo: () => {
          apply(before);
          this.#reproject(slideIndex);
        },
        redo: () => {
          apply(type);
          reseat();
          this.#reproject(slideIndex);
        },
      });
      this.#reproject(slideIndex);
      dialog.close();
    });
    this.#showModalDialog(dialog, () => (this.#layoutDialog = null));
    this.#layoutDialog = dialog;
  }

  /** Word's Reset: every placeholder child returns to the inherited
   *  position/size — the layout's placeholder geometry, falling back to the
   *  master's when the layout placeholder has no xfrm of its own. Text is
   *  untouched. */
  #resetSlideToLayout(): void {
    const presJson = this.#presJson;
    const slideIndex = this.#activeSlideIndex();
    const slide = presJson?.slides?.[slideIndex];
    if (!presJson || !slide) return;
    if (!slide.layout) return;
    const master =
      presJson.masters?.find((m) => m.name === (slide.master ?? m.name)) ?? presJson.masters?.[0];
    const layout = master?.layouts?.find(
      (candidate) => candidate.type === slide.layout || candidate.name === slide.layout,
    );
    const moved = resetSlidePlaceholders(slide, layout, master);
    if (moved.length === 0) return;
    this.#pushEdit({
      undo: () => {
        for (const item of moved) Object.assign(item.child.shape, item.before);
        this.#reproject(slideIndex);
      },
      redo: () => {
        for (const item of moved) Object.assign(item.child.shape, item.after);
        this.#reproject(slideIndex);
      },
    });
    this.#reproject(this.#activeSlideIndex());
  }

  /** Header/footer visibility instantiates real dt/ftr/sldNum placeholders on
   *  save; the model stays the single source, so children never duplicate it. */
  #openHeaderFooterDialog(): void {
    if (!this.#pres || this.#headerFooterDialog?.open) return;
    const slide = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!slide) return;
    const options = slide.headerFooter ?? {};
    const footerText = typeof options.footer === "string" ? options.footer : "";
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.header-footer.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body">
        <label class="hf-option"><input id="hf-date" type="checkbox" ${options.dateTime ? "checked" : ""}><span>${escapeHtml(t("ppt.header-footer.date", this))}</span></label>
        <label class="hf-option"><input id="hf-number" type="checkbox" ${options.slideNumber ? "checked" : ""}><span>${escapeHtml(t("ppt.header-footer.number", this))}</span></label>
        <label class="hf-option"><input id="hf-footer" type="checkbox" ${options.footer ? "checked" : ""}><span>${escapeHtml(t("ppt.header-footer.footer", this))}</span></label>
        <label class="dialog-field"><span>${escapeHtml(t("ppt.header-footer.text", this))}</span><input id="hf-text" value="${escapeHtml(footerText)}" placeholder="${escapeHtml(t("ppt.header-footer.placeholder", this))}"></label>
      </div>
      <div class="dialog-actions">
        <button data-hf-all>${escapeHtml(t("ppt.header-footer.apply-all", this))}</button>
        <button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button>
        <button data-hf-apply>${escapeHtml(t("ppt.header-footer.apply", this))}</button>
      </div>
    `;
    const read = (): { dateTime?: boolean; slideNumber?: boolean; footer?: string | boolean } => {
      const date = dialog.querySelector<HTMLInputElement>("#hf-date")?.checked ?? false;
      const number = dialog.querySelector<HTMLInputElement>("#hf-number")?.checked ?? false;
      const enabled = dialog.querySelector<HTMLInputElement>("#hf-footer")?.checked ?? false;
      const text = dialog.querySelector<HTMLInputElement>("#hf-text")?.value.trim() ?? "";
      return {
        ...(date ? { dateTime: true } : {}),
        ...(number ? { slideNumber: true } : {}),
        ...(enabled ? { footer: text || true } : {}),
      };
    };
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (target.closest("[data-hf-apply]")) {
        this.#setSlideHeaderFooter(read(), false);
        dialog.close();
      }
      if (target.closest("[data-hf-all]")) {
        this.#setSlideHeaderFooter(read(), true);
        dialog.close();
      }
    });
    this.#showModalDialog(dialog, () => (this.#headerFooterDialog = null));
    this.#headerFooterDialog = dialog;
  }

  #setSlideHeaderFooter(
    value: { dateTime?: boolean; slideNumber?: boolean; footer?: string | boolean },
    all: boolean,
  ): void {
    const presJson = this.#presJson;
    const at = this.#activeSlideIndex();
    const slides = presJson?.slides;
    if (!slides?.length) return;
    const targets = all ? slides.map((_, index) => index) : [at];
    const before = slides.map((slide) => structuredClone(slide.headerFooter));
    const apply = (snapshot: (typeof before)[number], index: number): void => {
      const slide = slides[index]!;
      if (snapshot) slide.headerFooter = structuredClone(snapshot);
      else delete slide.headerFooter;
    };
    before.forEach((snapshot, index) => apply(snapshot, index));
    targets.forEach((index) => apply(value, index));
    const finalAfter = slides.map((slide) => structuredClone(slide.headerFooter));
    this.#pushEdit({
      undo: () => before.forEach((snapshot, index) => apply(snapshot, index)),
      redo: () => finalAfter.forEach((snapshot, index) => apply(snapshot, index)),
    });
  }

  /** Shared native-dialog lifecycle: close removes the node and clears the
   *  host's handle, so HMR/undo cannot strand stale modal elements. */
  #showModalDialog(dialog: HTMLDialogElement, onClose: () => void): void {
    dialog.addEventListener("close", () => {
      onClose();
      dialog.remove();
    });
    this.shadowRoot?.append(dialog);
    dialog.showModal();
  }

  /** The active slide's comment list (the model persists through the real
   *  commentAuthors/comment parts on save). */
  #commentsOf(): { slide: SlideOptions; comments: SlideCommentOptions[] } | null {
    const slide = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!slide) return null;
    return { slide, comments: (slide.comments ??= []) };
  }

  /** A comment anchor at the selected object (or slide center), normalized to
   *  EMU — PowerPoint anchors comments to the point the reviewer inspected. */
  #commentAnchor(): { x: number; y: number } {
    const pres = this.#pres;
    const selected = this.#selection ? this.#selectedBox() : null;
    const box = selected ? this.#slideBoxOf(selected.box) : null;
    return {
      x: Math.round((box ? box.x + box.width / 2 : (pres?.widthPx ?? 720) / 2) * EMU_PER_PX),
      y: Math.round((box ? box.y + box.height / 2 : (pres?.heightPx ?? 405) / 2) * EMU_PER_PX),
    };
  }

  #openCommentsDialog(focusDraft = false): void {
    if (!this.#pres || this.#commentsDialog?.open) {
      if (focusDraft) this.#commentsDialog?.querySelector("textarea")?.focus();
      return;
    }
    const dialog = document.createElement("dialog");
    dialog.className = "comments-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.comments.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="comments-body">
        <label class="comment-draft"><textarea rows="3" placeholder="${escapeHtml(t("ppt.comments.placeholder", this))}"></textarea><button data-comment-post>${escapeHtml(t("ppt.comments.post", this))}</button></label>
        <div class="comment-list"></div>
      </div>
    `;
    dialog.querySelector(".comment-list")?.addEventListener("click", this.#onCommentsClick);
    dialog.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (target.closest("[data-comment-post]")) {
        const value = dialog.querySelector<HTMLTextAreaElement>("textarea")?.value.trim();
        if (value) this.#addSlideComment(value);
      }
    });
    this.#showModalDialog(dialog, () => (this.#commentsDialog = null));
    this.#commentsDialog = dialog;
    this.#renderCommentsDialog();
    if (focusDraft) dialog.querySelector("textarea")?.focus();
  }

  #toggleCommentsDialog(): void {
    if (this.#commentsDialog?.open) this.#commentsDialog.close();
    else this.#openCommentsDialog();
  }

  #renderCommentsDialog(): void {
    const dialog = this.#commentsDialog;
    const found = this.#commentsOf();
    const list = dialog?.querySelector<HTMLElement>(".comment-list");
    if (!dialog || !found || !list) return;
    if (found.comments.length === 0) {
      list.innerHTML = `<p class="comment-empty">${escapeHtml(t("ppt.comments.empty", this))}</p>`;
      return;
    }
    list.innerHTML = found.comments
      .map((comment, index) => {
        const author = escapeHtml(comment.author || t("ppt.comments.anonymous", this));
        const date = escapeHtml(comment.date ?? "");
        const text = escapeHtml(comment.text);
        return `<article class="comment" data-comment-index="${index}"><header><strong>${author}</strong><time>${date}</time></header><p>${text}</p><footer><button data-comment-edit>${escapeHtml(t("ppt.comments.edit", this))}</button><button data-comment-delete>${escapeHtml(t("ppt.comments.delete", this))}</button></footer></article>`;
      })
      .join("");
  }

  #addSlideComment(text: string): void {
    const found = this.#commentsOf();
    if (!found || !text.trim()) return;
    const author = this.getAttribute("user") || t("ppt.comments.anonymous", this);
    const initials =
      author
        .split(/\s+/u)
        .filter(Boolean)
        .slice(0, 2)
        .map((part) => part[0]!.toUpperCase())
        .join("") || "A";
    const comment: SlideCommentOptions = {
      author,
      initials,
      text,
      date: new Date().toLocaleString(),
      ...this.#commentAnchor(),
    };
    const comments = found.comments;
    comments.push(comment);
    this.#pushEdit({
      undo: () => {
        comments.pop();
        if (comments.length === 0) delete found.slide.comments;
        this.#renderCommentsDialog();
      },
      redo: () => {
        comments.push(comment);
        this.#renderCommentsDialog();
      },
    });
    const draft = this.#commentsDialog?.querySelector<HTMLTextAreaElement>("textarea");
    if (draft) draft.value = "";
    this.#renderCommentsDialog();
  }

  readonly #onCommentsClick = (event: Event): void => {
    const target = event.target as HTMLElement;
    const article = target.closest<HTMLElement>(".comment");
    const index = Number(article?.dataset.commentIndex ?? -1);
    const found = this.#commentsOf();
    if (!article || !found || index < 0 || index >= found.comments.length) return;
    if (target.closest("[data-comment-delete]")) {
      const before = structuredClone(found.comments);
      const [removed] = found.comments.splice(index, 1);
      if (!removed) return;
      if (found.comments.length === 0) delete found.slide.comments;
      const after = structuredClone(found.comments);
      this.#pushEdit({
        undo: () => {
          found.slide.comments = structuredClone(before);
          this.#renderCommentsDialog();
        },
        redo: () => {
          if (after.length === 0) delete found.slide.comments;
          else found.slide.comments = structuredClone(after);
          this.#renderCommentsDialog();
        },
      });
      this.#renderCommentsDialog();
      return;
    }
    if (target.closest("[data-comment-edit]")) {
      const comment = found.comments[index]!;
      const next = window.prompt(t("ppt.comments.prompt", this), comment.text)?.trim();
      if (next == null || next === comment.text) return;
      const before = comment.text;
      comment.text = next;
      comment.modified = true;
      this.#pushEdit({
        undo: () => {
          comment.text = before;
          if (!before) delete comment.modified;
          this.#renderCommentsDialog();
        },
        redo: () => {
          comment.text = next;
          comment.modified = true;
          this.#renderCommentsDialog();
        },
      });
      this.#renderCommentsDialog();
    }
  };

  /** The Shapes gallery's pick arms the drawer; the value is the prstGeom
   *  token (the main button draws the default rectangle). */
  #armShape(geometry?: string): void {
    this.#shapeDrawer?.arm(geometry ?? "rect");
  }

  /** Draw-tab tools own the surface cursor: select restores the normal
   *  gestures, eraser removes top-level objects under a click or drag. */
  #armDrawTool(tool: "select" | "eraser" | "pen"): void {
    this.#drawTool = tool;
    this.#shapeDrawer?.disarm();
    this.#stage.style.cursor = tool === "eraser" ? "cell" : tool === "pen" ? "crosshair" : "";
  }

  /** Pen captures a literal point sequence and commits one OOXML custGeom
   *  shape; the preview is a throwaway SVG over the canvas, while the model
   *  only sees the real vector freeform at pointer-up. */
  #startPen(event: PointerEvent): void {
    const pres = this.#pres;
    const start = this.#stagePointOf(event);
    if (!pres || !start) return;
    event.preventDefault();
    const points = [{ x: start.x, y: start.y }];
    const host = this.#canvasHost();
    const stageRect = this.#stage.getBoundingClientRect();
    const hostRect = host.getBoundingClientRect();
    const scale = stageRect.width / pres.widthPx;
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    svg.append(path);
    Object.assign(svg.style, {
      position: "absolute",
      left: `${stageRect.left - hostRect.left}px`,
      top: `${stageRect.top - hostRect.top + (start.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX) * scale}px`,
      width: `${pres.widthPx * scale}px`,
      height: `${pres.heightPx * scale}px`,
      overflow: "visible",
      pointerEvents: "none",
      zIndex: "30",
    } satisfies Partial<CSSStyleDeclaration>);
    Object.assign(path.style, {
      fill: "none",
      stroke: "#262626",
      strokeWidth: `${1.333 * scale}px`,
      strokeLinecap: "round",
      strokeLinejoin: "round",
    } satisfies Partial<CSSStyleDeclaration>);
    host.append(svg);
    const paint = (): void => {
      path.setAttribute(
        "d",
        points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "),
      );
    };
    paint();
    const onMove = (move: PointerEvent): void => {
      const next = this.#stagePointOf(move);
      if (!next || next.slide !== start.slide) return;
      points.push({ x: next.x, y: next.y });
      paint();
    };
    const stop = (): void => {
      document.removeEventListener("pointermove", onMove, { capture: true });
      document.removeEventListener("pointerup", stop, { capture: true });
      document.removeEventListener("pointercancel", stop, { capture: true });
      svg.remove();
      const stroke = makePenStroke(points);
      if (stroke) this.#insertChild(stroke, start.slide);
    };
    document.addEventListener("pointermove", onMove, { capture: true });
    document.addEventListener("pointerup", stop, { capture: true });
    document.addEventListener("pointercancel", stop, { capture: true });
  }

  /** Erase every top-level object the drag touches; each removal records its
   *  own undo step, matching PowerPoint's stroke-wise object erasure. */
  #startEraser(event: PointerEvent): void {
    const presJson = this.#presJson;
    if (!presJson) return;
    const point = this.#stagePointOf(event);
    if (!point) return;
    event.preventDefault();
    const eraseAt = (x: number, y: number): void => {
      const slide = presJson.slides?.[point.slide];
      if (!slide) return;
      const hit = hitSlide(slideHits(slide), x, y);
      if (hit < 0) return;
      const selection = this.#selection;
      const children = slide.children;
      const child = children?.[hit];
      if (!children || !child) return;
      children.splice(hit, 1);
      if (selection?.slide === point.slide && selection.child === hit) this.#select(null);
      this.#pushEdit({
        undo: () => {
          children.splice(hit, 0, child);
          this.#reproject(point.slide);
        },
        redo: () => {
          children.splice(hit, 1);
          this.#reproject(point.slide);
        },
      });
      this.#reproject(point.slide);
    };
    eraseAt(point.x, point.y);
    const onMove = (move: PointerEvent): void => {
      const next = this.#stagePointOf(move);
      if (!next || next.slide !== point.slide) return;
      eraseAt(next.x, next.y);
    };
    const stop = (): void => {
      document.removeEventListener("pointermove", onMove, { capture: true });
      document.removeEventListener("pointerup", stop, { capture: true });
      document.removeEventListener("pointercancel", stop, { capture: true });
    };
    document.addEventListener("pointermove", onMove, { capture: true });
    document.addEventListener("pointerup", stop, { capture: true });
    document.addEventListener("pointercancel", stop, { capture: true });
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

  // ── Slide Sorter ────────────────────────────────────────────────────────

  /** Reorder deck slides as one reversible structural edit. Selection and
   *  panes reset because every child index after the removal point moves. */
  #moveSlide(from: number, to: number): void {
    const slides = this.#presJson?.slides;
    if (
      !slides ||
      from === to ||
      from < 0 ||
      to < 0 ||
      from >= slides.length ||
      to >= slides.length
    )
      return;
    const before = slides.map((slide) => structuredClone(slide));
    const [moved] = slides.splice(from, 1);
    if (!moved) return;
    slides.splice(to, 0, moved);
    const after = slides.map((slide) => structuredClone(slide));
    const restore = (value: SlideOptions[]): void => {
      slides.splice(0, slides.length, ...structuredClone(value));
      this.#select(null);
      this.#reproject();
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
  }

  /** Open Slide Sorter over cropped thumbnails of the existing Leafer
   *  paint. The thumbnails already share the deck's render loop, so this
   *  avoids a second painter while staying visually faithful. */
  #toggleSlideSorter(): void {
    if (this.#slideSorter?.open) return this.#slideSorter.close();
    const pres = this.#pres;
    if (!pres || pres.slides.length === 0) return;
    const source = this.shadowRoot?.querySelector<HTMLCanvasElement>(".thumb-stage canvas");
    if (!source) return;
    const dialog = document.createElement("dialog");
    dialog.className = "slide-sorter";
    dialog.innerHTML = `<div class="sorter-head"><strong>${escapeHtml(t("ppt.sorter.title", this))}</strong><button data-sorter-close>×</button></div><div class="sorter-grid"></div>`;
    const grid = dialog.querySelector<HTMLElement>(".sorter-grid")!;
    const ratio = source.width / Math.max(1, source.clientWidth);
    const sourceSlide = pres.heightPx * this.#thumbScale * ratio;
    const sourceGap = THUMB_GAP_PX * ratio;
    const width = 240;
    for (let index = 0; index < pres.slides.length; index += 1) {
      const card = document.createElement("figure");
      card.className = "sorter-card";
      card.draggable = true;
      card.dataset.sortIndex = String(index);
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = Math.round(width * (pres.heightPx / pres.widthPx));
      canvas
        .getContext("2d")
        ?.drawImage(
          source,
          0,
          index * (sourceSlide + sourceGap),
          source.width,
          sourceSlide,
          0,
          0,
          canvas.width,
          canvas.height,
        );
      const label = document.createElement("figcaption");
      label.textContent = String(index + 1);
      card.append(canvas, label);
      grid.append(card);
    }
    dialog.addEventListener("dragstart", (event) => {
      const card = (event.target as HTMLElement).closest<HTMLElement>("[data-sort-index]");
      this.#sorterDragIndex = card ? Number(card.dataset.sortIndex) : -1;
      this.#sorterDragged = false;
    });
    dialog.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer!.dropEffect = "move";
    });
    dialog.addEventListener("drop", (event) => {
      event.preventDefault();
      const target = (event.target as HTMLElement).closest<HTMLElement>("[data-sort-index]");
      const to = target ? Number(target.dataset.sortIndex) : -1;
      if (this.#sorterDragIndex >= 0 && to >= 0 && this.#sorterDragIndex !== to) {
        this.#sorterDragged = true;
        this.#moveSlide(this.#sorterDragIndex, to);
        this.#renderSlideSorter();
        return;
      }
      this.#renderSlideSorter();
    });
    dialog.addEventListener("click", (event) => {
      if ((event.target as HTMLElement).closest("[data-sorter-close]")) return dialog.close();
      const card = (event.target as HTMLElement).closest<HTMLElement>("[data-sort-index]");
      if (!card || this.#sorterDragged) return;
      this.#sorterDragged = false;
      dialog.close();
      this.#revealSlide(Number(card.dataset.sortIndex));
    });
    dialog.addEventListener("close", () => {
      this.#slideSorter = null;
      dialog.remove();
    });
    this.shadowRoot?.append(dialog);
    this.#slideSorter = dialog;
    dialog.showModal();
  }

  /** Refresh an open sorter after an undo/redo or a drag reorder. */
  #renderSlideSorter(): void {
    if (this.#slideSorter?.open) {
      this.#slideSorter.close();
      requestAnimationFrame(() => this.#toggleSlideSorter());
    }
  }

  /** A western word candidate, mirroring the DOCX checker's first pass:
   *  CJK, numbers, URLs and symbols are outside the built-in dictionary. */
  #isSpellCandidate(word: string): boolean {
    return /^[A-Za-z][A-Za-z'’-]*$/.test(word);
  }

  /** Scan editable shape, table-cell and notes text. Each issue keeps its
   *  live model source so replacement reuses the editor's text writer. */
  #findSpellingIssues(): PresentationSpellingIssue[] {
    const presJson = this.#presJson;
    if (!presJson?.slides) return [];
    const segmenter = new Intl.Segmenter("en", { granularity: "word" });
    const issues: PresentationSpellingIssue[] = [];
    const addText = (
      slide: number,
      text: string,
      source: PresentationSpellingIssue["source"],
    ): void => {
      for (const part of segmenter.segment(text)) {
        const word = part.segment;
        if (!part.isWordLike || !this.#isSpellCandidate(word)) continue;
        const key = word.toLowerCase();
        if (englishWords.has(key) || englishWords.has(word)) continue;
        issues.push({ slide, word, from: part.index, to: part.index + word.length, source });
      }
    };
    const visit = (slide: number, child: SlideChild): void => {
      if ("shape" in child) {
        addText(slide, textOf({ kind: "shape", child: child.shape }), {
          kind: "shape",
          child: child.shape,
        });
      }
      if ("table" in child) {
        for (const origin of tableGridOf(child.table).origins) {
          addText(slide, textOf({ kind: "cell", cell: origin.cell }), {
            kind: "cell",
            cell: origin.cell,
          });
        }
      }
      if ("group" in child) for (const member of child.group.children ?? []) visit(slide, member);
    };
    presJson.slides.forEach((slide, index) => {
      for (const child of slide.children ?? []) visit(index, child);
      addText(index, slideNotesOf(slide), { kind: "notes", slide });
    });
    return issues;
  }

  /** Open/re-render Spelling. The dialog stays open when Ignore replaces the
   *  active issue; it closes automatically when the deck is clean. */
  #openSpellCheck(): void {
    if (this.#spellDialog?.open) return this.#renderSpellDialog();
    if (!this.#presJson) return;
    this.#spellIssues = this.#findSpellingIssues();
    this.#spellAt = Math.min(this.#spellAt, Math.max(0, this.#spellIssues.length - 1));
    const dialog = document.createElement("dialog");
    dialog.className = "spell-dialog";
    dialog.addEventListener("click", (event) => {
      const action = (event.target as HTMLElement).closest<HTMLElement>("[data-spell-action]");
      if (!action) return;
      const value = action.dataset.spellValue;
      const kind = action.dataset.spellAction;
      if (kind === "close") dialog.close();
      else if (kind === "replace") this.#replaceSpellingIssue(value ?? "");
      else if (kind === "ignore-once") this.#skipSpellingIssue(false);
      else if (kind === "ignore-all") this.#skipSpellingIssue(true);
      else if (kind === "add") addSpellWord(this.#spellIssues[this.#spellAt]?.word ?? "");
      else if (kind === "nav")
        this.#spellAt =
          (this.#spellAt + Number(value) + this.#spellIssues.length) %
          Math.max(1, this.#spellIssues.length);
      this.#renderSpellDialog();
    });
    dialog.addEventListener("close", () => {
      this.#spellDialog = null;
      dialog.remove();
    });
    this.shadowRoot?.append(dialog);
    this.#spellDialog = dialog;
    this.#renderSpellDialog();
    dialog.showModal();
  }

  #renderSpellDialog(): void {
    const dialog = this.#spellDialog;
    if (!dialog) return;
    const issue = this.#spellIssues[this.#spellAt];
    const suggestions = issue ? spellSuggestions(issue.word) : [];
    dialog.innerHTML = `
      <div class="spell-head"><strong>${escapeHtml(t("ppt.spell.title", this))}</strong><button data-spell-action="close">×</button></div>
      <div class="spell-body">
        ${
          issue
            ? `<span class="spell-counter">${escapeHtml(t("ppt.spell.counter", this))}</span><span class="spell-word">${escapeHtml(issue.word)}</span><span class="spell-section">${escapeHtml(t("ppt.spell.suggestions", this))}</span><div class="spell-suggestions">${suggestions.map((s) => `<button data-spell-action="replace" data-spell-value="${escapeHtml(s)}">${escapeHtml(s)}</button>`).join("") || `<span>${escapeHtml(t("ppt.spell.none", this))}</span>`}</div><div class="spell-actions"><button data-spell-action="ignore-once">${escapeHtml(t("ppt.spell.ignore-once", this))}</button><button data-spell-action="ignore-all">${escapeHtml(t("ppt.spell.ignore-all", this))}</button><button data-spell-action="add">${escapeHtml(t("ppt.spell.add", this))}</button></div>`
            : `<div class="spell-clean">${escapeHtml(t("ppt.spell.clean", this))}</div>`
        }
      </div>
      <div class="spell-nav"><button data-spell-action="nav" data-spell-value="-1">${escapeHtml(t("ppt.spell.previous", this))}</button><button data-spell-action="nav" data-spell-value="1">${escapeHtml(t("ppt.spell.next", this))}</button></div>
    `;
    dialog
      .querySelector<HTMLElement>(".spell-counter")
      ?.replaceChildren(`${this.#spellAt + 1} / ${this.#spellIssues.length}`);
  }

  /** Replace only the active occurrence. The model write is undoable; the
   *  issue list is rebuilt after projection so positions stay authoritative. */
  #replaceSpellingIssue(replacement: string): void {
    const issue = this.#spellIssues[this.#spellAt];
    if (!issue) return;
    const source = issue.source;
    if (source.kind === "notes") {
      const before = slideNotesOf(source.slide);
      const after = before.slice(0, issue.from) + replacement + before.slice(issue.to);
      writeSlideNotes(source.slide, after);
      this.#pushEdit({
        undo: () => writeSlideNotes(source.slide, before),
        redo: () => writeSlideNotes(source.slide, after),
      });
      this.#syncNotesPane(issue.slide);
    } else {
      const beforeText = textOf(source);
      const afterText = beforeText.slice(0, issue.from) + replacement + beforeText.slice(issue.to);
      const model = source.kind === "shape" ? source.child : source.cell;
      const before = structuredClone(model);
      writeText(source, afterText);
      const after = structuredClone(model);
      const restore = (value: typeof before): void => {
        for (const key of Object.keys(model) as (keyof typeof model)[]) delete model[key];
        Object.assign(model, value);
      };
      this.#pushEdit({
        undo: () => {
          restore(before);
          this.#reproject(issue.slide);
        },
        redo: () => {
          restore(after);
          this.#reproject(issue.slide);
        },
      });
      this.#reproject(issue.slide);
    }
    this.#spellIssues.splice(this.#spellAt, 1);
    this.#spellAt = Math.min(this.#spellAt, Math.max(0, this.#spellIssues.length - 1));
  }

  /** Ignore Once keeps other occurrences active; Ignore All removes every
   *  case-insensitive occurrence in this session via the shared dictionary. */
  #skipSpellingIssue(all: boolean): void {
    const issue = this.#spellIssues[this.#spellAt];
    if (!issue) return;
    if (all) {
      ignoreSpellWord(issue.word);
      this.#spellIssues = this.#spellIssues.filter(
        (candidate) => candidate.word.toLowerCase() !== issue.word.toLowerCase(),
      );
    } else {
      this.#spellIssues.splice(this.#spellAt, 1);
    }
    this.#spellAt = Math.min(this.#spellAt, Math.max(0, this.#spellIssues.length - 1));
  }

  // ── Transitions ──────────────────────────────────────────────────────────

  /** The gallery's pick on the active slide: the effect token lands as the
   *  slide's transition; none clears it (playback-only — nothing paints). */
  #setTransition(value: string): void {
    const host = this.#presJson?.slides?.[this.#activeSlideIndex()];
    if (!host) return;
    const before = host.transition;
    if (value !== "none" && !TRANSITION_PRESETS.has(value as TransitionType)) return;
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

  /** An Animation split's pick on the selected object. Animate replaces the
   *  shape's effects; Add Animation appends another one, matching PowerPoint's
   *  two commands. Entries address the shape by its cNvPr name (an unnamed
   *  shape takes one unique among its siblings — the file format resolves
   *  names to spTgt ids at compile time). */
  #applyAnimation(value?: string, mode: "replace" | "add" = "replace"): void {
    const preset = value ?? "fade";
    if (preset !== "none" && !ANIMATION_PRESETS.has(preset as SlideAnimation["type"])) return;
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
    const other = entries.filter((entry) => entry.shapeName !== name);
    const mine = entries.filter((entry) => entry.shapeName === name);
    let after: SlideAnimation[] | undefined;
    if (mode === "add" && preset === "none") {
      const keptMine = mine.slice(0, -1);
      after = [...other, ...keptMine];
    } else if (mode === "add") {
      after = [...entries, { type: preset, shapeName: name } as SlideAnimation];
    } else {
      after =
        preset === "none" ? other : [...other, { type: preset, shapeName: name } as SlideAnimation];
    }
    if (after.length === 0) after = undefined;
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const restore = (value: SlideAnimation[] | string | undefined): void => {
      if (value === undefined) delete host.animations;
      else host.animations = structuredClone(value);
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
    if (this.#rightPane === "animation") this.#renderAnimationPane();
  }

  /** Animation Pane timing controls: PowerPoint's Start / Duration / Delay
   *  fields write the same official p:cTn attributes the compiler emits. */
  #updateAnimationSetting(index: number, setting: string, value: string): void {
    const host = this.#presJson?.slides?.[this.#activeSlideIndex()];
    const entries = Array.isArray(host?.animations) ? host!.animations! : [];
    const entry = entries[index];
    if (!host || !entry || typeof entry !== "object") return;
    const before = structuredClone(entry);
    const after = structuredClone(entry);
    if (setting === "trigger") {
      if (value !== "onClick" && value !== "withPrevious" && value !== "afterPrevious") return;
      after.trigger = value;
    } else {
      const seconds = Number(value);
      if (!Number.isFinite(seconds) || seconds < 0) return;
      const milliseconds = Math.round(seconds * 1000);
      if (setting === "duration") after.duration = milliseconds;
      else if (setting === "delay") after.delay = milliseconds;
      else return;
    }
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const restore = (next: SlideAnimation): void => {
      entries[index] = structuredClone(next);
      this.#renderAnimationPane();
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
  }

  // ── Slide show ───────────────────────────────────────────────────────────

  /** PowerPoint's Set Up Show. The range and playback toggles govern the
   *  live show; settings stay in the editor session until PPTX presentation
   *  properties are exposed by the format package. */
  #openShowSetup(): void {
    const pres = this.#pres;
    if (!pres || this.#showSetupDialog?.open) return;
    const dialog = document.createElement("dialog");
    dialog.className = "show-setup";
    const last = pres.slides.length;
    const setup = this.#showSetup;
    const from = Math.min(setup.from, last);
    const to = Math.min(Number.isFinite(setup.to) ? setup.to : last, last);
    dialog.innerHTML = `
      <div class="setup-head"><strong>${escapeHtml(t("ppt.setup.title", this))}</strong><button data-setup-action="cancel">×</button></div>
      <div class="setup-body">
        <fieldset><legend>${escapeHtml(t("ppt.setup.slides", this))}</legend>
          <label><input type="radio" name="setup-range" value="all" ${setup.to === Number.POSITIVE_INFINITY ? "checked" : ""}> ${escapeHtml(t("ppt.setup.all", this))}</label>
          <label class="range"><input type="radio" name="setup-range" value="range" ${Number.isFinite(setup.to) ? "checked" : ""}> <input id="setup-from" type="number" min="1" max="${last}" value="${from}"> <span>–</span> <input id="setup-to" type="number" min="1" max="${last}" value="${to}"></label>
        </fieldset>
        <fieldset><legend>${escapeHtml(t("ppt.setup.options", this))}</legend>
          <label><input id="setup-loop" type="checkbox" ${setup.loop ? "checked" : ""}> ${escapeHtml(t("ppt.setup.loop", this))}</label>
          <label><input id="setup-narration" type="checkbox" ${setup.narration ? "checked" : ""}> ${escapeHtml(t("ppt.setup.narration", this))}</label>
          <label><input id="setup-animations" type="checkbox" ${setup.animations ? "checked" : ""}> ${escapeHtml(t("ppt.setup.animations", this))}</label>
          <label><input id="setup-timings" type="checkbox" ${setup.timings ? "checked" : ""}> ${escapeHtml(t("ppt.setup.timings", this))}</label>
        </fieldset>
      </div>
      <div class="setup-actions"><button data-setup-action="cancel">${escapeHtml(t("ppt.setup.cancel", this))}</button><button data-setup-action="ok">${escapeHtml(t("ppt.setup.ok", this))}</button></div>
    `;
    dialog.addEventListener("click", (event) => {
      const action = (event.target as HTMLElement).closest<HTMLElement>("[data-setup-action]");
      if (action?.dataset.setupAction === "ok") {
        const range = dialog.querySelector<HTMLInputElement>(
          'input[name="setup-range"]:checked',
        )?.value;
        const fromValue = Number(dialog.querySelector<HTMLInputElement>("#setup-from")?.value);
        const toValue = Number(dialog.querySelector<HTMLInputElement>("#setup-to")?.value);
        this.#showSetup = {
          from: Math.max(1, Math.min(last, Number.isFinite(fromValue) ? fromValue : 1)),
          to:
            range === "range"
              ? Math.max(1, Math.min(last, Number.isFinite(toValue) ? toValue : last))
              : Number.POSITIVE_INFINITY,
          loop: dialog.querySelector<HTMLInputElement>("#setup-loop")?.checked === true,
          narration: dialog.querySelector<HTMLInputElement>("#setup-narration")?.checked !== true,
          animations: dialog.querySelector<HTMLInputElement>("#setup-animations")?.checked !== true,
          timings: dialog.querySelector<HTMLInputElement>("#setup-timings")?.checked !== false,
        };
        if (this.#showSetup.to < this.#showSetup.from) this.#showSetup.to = this.#showSetup.from;
        dialog.close();
      } else if (action?.dataset.setupAction === "cancel") dialog.close();
    });
    dialog.addEventListener("close", () => {
      this.#showSetupDialog = null;
      dialog.remove();
    });
    this.shadowRoot?.append(dialog);
    this.#showSetupDialog = dialog;
    dialog.showModal();
  }

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
    this.#gotoShowSlide(from === "beginning" ? this.#showSetup.from - 1 : this.#activeSlideIndex());
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
    const setupFrom = Math.max(0, this.#showSetup.from - 1);
    const setupTo = Math.min(pres.slides.length - 1, this.#showSetup.to - 1);
    this.#showingSlide = Math.max(setupFrom, Math.min(setupTo, index));
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
    if (
      !this.#showSetup.animations ||
      !slide ||
      !slideGroup ||
      !Array.isArray(entries) ||
      entries.length === 0
    )
      return;
    const hits = slideHits(slide);
    const inBox = (box: Box, x: number, y: number): boolean =>
      x >= box.x - 1 && x <= box.x + box.width + 1 && y >= box.y - 1 && y <= box.y + box.height + 1;
    let clickCursor = 0;
    let lastStart = 0;
    let lastEnd = 0;
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
      const trigger = entry.trigger ?? "onClick";
      const base =
        trigger === "withPrevious"
          ? lastStart
          : trigger === "afterPrevious"
            ? lastEnd
            : clickCursor;
      const startAt = base + (entry.delay ?? 0);
      const options = { duration, delay: startAt, jump: true, easing: "ease-out" as const };
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
      clickCursor = Math.max(clickCursor, startAt + duration);
      lastStart = startAt;
      lastEnd = startAt + duration;
    }
  }

  /** Advance within the configured range; `loop` wraps back to its first
   *  slide, otherwise the show stays on the last slide. */
  #advanceShow(): void {
    const pres = this.#pres;
    if (!pres) return;
    const last = Math.min(pres.slides.length - 1, this.#showSetup.to - 1);
    if (this.#showingSlide >= last && this.#showSetup.loop)
      this.#gotoShowSlide(Math.max(0, this.#showSetup.from - 1));
    else this.#gotoShowSlide(this.#showingSlide + 1);
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
    const show = pane.hidden || this.#rightPane !== "select";
    this.#rightPane = show ? "select" : null;
    pane.hidden = !show;
    if (show) this.#renderSelectionPane();
  }

  /** PowerPoint's Animation Pane: the active slide's effects in play order,
   * with target selection and order/delete operations. */
  #toggleAnimationPane(): void {
    const pane = this.#selectionPane();
    if (!pane) return;
    const show = pane.hidden || this.#rightPane !== "animation";
    this.#rightPane = show ? "animation" : null;
    pane.hidden = !show;
    if (show) this.#renderAnimationPane();
  }

  /** The selection pane's row label: the cNvPr name when the child carries
   *  one, else the localized kind plus its position. */
  #childLabel(
    child: SlideChild,
    index: number,
    path: readonly number[] = [],
  ): { label: string; hidden: boolean } {
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
                        : "ole" in child
                          ? "object"
                          : "table";
    const name = nv?.name?.trim();
    const position = path.length > 0 ? path.at(-1)! + 1 : index + 1;
    return {
      label: name || `${t(`ppt.select.${kind}`, this)} ${position}`,
      hidden: nv?.hidden === true,
    };
  }

  /** Stamp the active slide's object list. A no-op while the pane is hidden —
   *  every refresh trigger re-runs this when it shows. */
  #renderSelectionPane(): void {
    const pane = this.#selectionPane();
    const list = this.selectList;
    if (!pane || pane.hidden || !list) return;
    if (this.#rightPane === "animation") return this.#renderAnimationPane();
    const head = pane.querySelector<HTMLElement>(".select-head");
    if (head) head.textContent = t("ppt.select.title", this);
    const slide = this.#activeSlideIndex();
    const children = this.#presJson?.slides?.[slide]?.children ?? [];
    const row = (child: SlideChild, root: number, path: readonly number[]): string => {
      const { label, hidden } = this.#childLabel(child, root, path);
      const member = path.join(".");
      const active =
        this.#selection?.slide === slide &&
        this.#selection.child === root &&
        pathsEqual(this.#selection.member ?? [], path);
      const indent = `style="padding-inline-start:${16 + path.length * 18}px"`;
      const memberAttr = path.length > 0 ? ` data-member="${member}"` : "";
      const html = `<div class="select-row" data-child="${root}"${memberAttr} data-hidden="${hidden}"${active ? ' data-active="true"' : ""}${indent}><button class="select-eye" data-eye="${root}"${memberAttr}>${hidden ? "○" : "●"}</button><span class="select-name">${escapeHtml(label)}</span></div>`;
      return "group" in child
        ? html +
            (child.group.children ?? [])
              .map((item, index) => row(item, root, [...path, index]))
              .join("")
        : html;
    };
    list.innerHTML = children.map((child, index) => row(child, index, [])).join("");
  }

  /** One pane row per animation entry: target label, localized effect, and
   *  the compact move/delete controls PowerPoint shows on hover. */
  #renderAnimationPane(): void {
    const pane = this.#selectionPane();
    const list = this.selectList;
    if (!pane || pane.hidden || !list) return;
    const head = pane.querySelector<HTMLElement>(".select-head");
    if (head) head.textContent = t("ppt.animation.title", this);
    const slide = this.#activeSlideIndex();
    const host = this.#presJson?.slides?.[slide];
    const children = host?.children ?? [];
    const entries = Array.isArray(host?.animations) ? host!.animations! : [];
    list.innerHTML = entries
      .map((entry, index) => {
        const target = children.findIndex(
          (child) => nonVisualOf(child)?.name?.trim() === entry.shapeName?.trim(),
        );
        const label =
          target >= 0
            ? this.#childLabel(children[target]!, target).label
            : (entry.shapeName ?? t("ppt.animation.unnamed", this));
        const effect = t(`ppt.ribbon.animate.${entry.type}`, this);
        const active =
          target >= 0 && this.#selection?.slide === slide && this.#selection.child === target;
        const trigger = entry.trigger ?? "onClick";
        const options: readonly SlideAnimation["trigger"][] = [
          "onClick",
          "withPrevious",
          "afterPrevious",
        ];
        return `<div class="select-row" data-child="${target}" data-animation="${index}"${active ? ' data-active="true"' : ""}>
          <span class="select-name"><strong>${escapeHtml(label)}</strong><br>${escapeHtml(effect)}</span>
          <div class="animation-controls">
            <label><span>${escapeHtml(t("ppt.animation.start", this))}</span><select data-animation-setting="trigger" data-animation-index="${index}">${options.map((option) => `<option value="${option}"${option === trigger ? " selected" : ""}>${escapeHtml(t(`ppt.animation.${option}`, this))}</option>`).join("")}</select></label>
            <label><span>${escapeHtml(t("ppt.animation.duration", this))}</span><input type="number" min="0" step="0.1" value="${((entry.duration ?? 500) / 1000).toFixed(1)}" data-animation-setting="duration" data-animation-index="${index}"></label>
            <label><span>${escapeHtml(t("ppt.animation.delay", this))}</span><input type="number" min="0" step="0.1" value="${((entry.delay ?? 0) / 1000).toFixed(1)}" data-animation-setting="delay" data-animation-index="${index}"></label>
          </div>
          <span class="pane-actions"><button data-animation-move="${index}" data-direction="-1" title="${escapeHtml(t("ppt.animation.move-up", this))}">↑</button><button data-animation-move="${index}" data-direction="1" title="${escapeHtml(t("ppt.animation.move-down", this))}">↓</button><button data-animation-delete="${index}" title="${escapeHtml(t("ppt.animation.delete", this))}">×</button></span>
        </div>`;
      })
      .join("");
  }

  /** Reorder one animation as a reversible slide-level edit. PowerPoint's
   *  arrows move the highlighted entry; no-op at either end. */
  #moveAnimation(index: number, delta: number): void {
    const slideIndex = this.#activeSlideIndex();
    const host = this.#presJson?.slides?.[slideIndex];
    const entries = host?.animations;
    const to = index + delta;
    if (!host || !Array.isArray(entries) || to < 0 || to >= entries.length) return;
    const before = structuredClone(entries);
    const after = structuredClone(entries);
    [after[index], after[to]] = [after[to]!, after[index]!];
    const restore = (value: SlideAnimation[]): void => {
      host.animations = structuredClone(value);
      this.#renderAnimationPane();
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
  }

  /** Remove one animation entry; a named target stays in the deck because
   *  its cNvPr may also be referenced by other effects or add-ins. */
  #deleteAnimation(index: number): void {
    const slideIndex = this.#activeSlideIndex();
    const host = this.#presJson?.slides?.[slideIndex];
    const entries = host?.animations;
    if (!host || !Array.isArray(entries) || index < 0 || index >= entries.length) return;
    const before = structuredClone(entries);
    const after = before.filter((_, i) => i !== index);
    const restore = (value: SlideAnimation[]): void => {
      if (value.length === 0) delete host.animations;
      else host.animations = structuredClone(value);
      this.#renderAnimationPane();
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    restore(after);
  }

  readonly #onSelectListClick = (event: Event): void => {
    const target = event.target as HTMLElement;
    const move = target.closest<HTMLElement>("[data-animation-move]");
    if (move)
      return this.#moveAnimation(
        Number(move.dataset.animationMove),
        Number(move.dataset.direction),
      );
    const remove = target.closest<HTMLElement>("[data-animation-delete]");
    if (remove) return this.#deleteAnimation(Number(remove.dataset.animationDelete));
    const eye = target.closest<HTMLElement>("[data-eye]");
    if (eye) return this.#toggleChildHidden(Number(eye.dataset.eye), this.#pathOf(eye));
    const row = target.closest<HTMLElement>("[data-child]");
    if (!row) return;
    const slide = this.#activeSlideIndex();
    this.#select({
      slide,
      child: Number(row.dataset.child),
      ...(row.dataset.member ? { member: this.#pathOf(row)! } : {}),
    });
    this.#revealSlide(slide);
  };

  /** A selection-pane path (`1.2`) as a fresh immutable member path. */
  #pathOf(element: HTMLElement): number[] | undefined {
    const value = element.dataset.member;
    return value ? value.split(".").map(Number) : undefined;
  }

  /** The eye toggle: cNvPr @hidden on the child (tables have no surface and
   *  stay put) as one reversible edit — hidden objects drop out of the paint. */
  #toggleChildHidden(index: number, path?: readonly number[]): void {
    const slide = this.#activeSlideIndex();
    const root = this.#presJson?.slides?.[slide]?.children?.[index];
    const child = root ? this.#childAt(root, path) : undefined;
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

  /** Every source cell under the current grid selection (all origins when no
   *  block is active), so Table Design commands target PowerPoint's active
   *  cell range rather than repainting an unrelated part of the table. */
  #selectedTableCells(): { table: TableOptions; cells: TableCellOptions[] } | null {
    const found = this.#tableMemberOf();
    if (!found) return null;
    const origins = tableGridOf(found.table).origins;
    const cells = this.#tableSelection
      ? tableSelectionCells(found.member, this.#tableSelection)
          .map(
            ({ row, col }) =>
              origins.find((origin) => origin.row === row && origin.col === col)?.cell,
          )
          .filter((cell): cell is TableCellOptions => Boolean(cell))
      : (() => {
          const active = this.#editedCell();
          return active ? [active] : origins.map((origin) => origin.cell);
        })();
    return cells.length > 0 ? { table: found.table, cells } : null;
  }

  /** Apply one reversible mutation to the selected table's official model. */
  #mutateSelectedTable(apply: (table: TableOptions) => void): void {
    const found = this.#tableMemberOf();
    if (!found) return;
    const table = found.table;
    const before = structuredClone(table);
    apply(table);
    const after = structuredClone(table);
    const restore = (value: TableOptions): void => {
      for (const key of Object.keys(table) as (keyof TableOptions)[]) delete table[key];
      Object.assign(table, value);
      this.#reproject(found.slide);
    };
    this.#pushEdit({ undo: () => restore(before), redo: () => restore(after) });
    this.#reproject(found.slide);
    this.#syncTableCellSizeControls();
  }

  /** PowerPoint's Table Style Options: the six official tblLook flags. */
  #toggleTableLook(value?: string): void {
    if (
      value !== "firstRow" &&
      value !== "lastRow" &&
      value !== "firstCol" &&
      value !== "lastCol" &&
      value !== "bandRow" &&
      value !== "bandCol"
    )
      return;
    this.#mutateSelectedTable((table) => {
      if (table[value] === true) delete table[value];
      else table[value] = true;
    });
  }

  /** A built-in style reference; an inline custom style is retired because it
   *  would outrank the selected PowerPoint style on the next projection. */
  #setTableStyle(value?: string): void {
    if (!value || !TABLE_STYLE_IDS.has(value)) return;
    this.#mutateSelectedTable((table) => {
      table.tableStyleId = value;
      delete table.tableStyle;
    });
  }

  /** Cell shading accepts the color picker's six-digit value; `none` is the
   *  official way to remove the direct fill. */
  #setTableCellShading(value?: string): void {
    const selected = this.#selectedTableCells();
    if (!selected) return;
    this.#mutateSelectedTable((table) => {
      const origins = new Set(tableGridOf(table).origins.map((origin) => origin.cell));
      for (const cell of selected.cells) {
        if (!origins.has(cell)) continue;
        if (value?.toLowerCase() === "none") delete cell.fill;
        else if (value && /^[0-9a-f]{6}$/i.test(value)) cell.fill = value.toUpperCase();
      }
    });
  }

  /** Ribbon border picks write the same official cell edge model the canvas
   *  paints; outside borders stay on the table frame, sides target cells. */
  #setTableCellBorders(value?: string): void {
    if (
      value !== "all" &&
      value !== "outside" &&
      value !== "none" &&
      value !== "top" &&
      value !== "bottom" &&
      value !== "left" &&
      value !== "right"
    )
      return;
    this.#mutateSelectedTable((table) => {
      if (value === "outside") {
        table.borders = { top: {}, right: {}, bottom: {}, left: {} };
        return;
      }
      const clearCell = (cell: TableCellOptions): void => {
        if (!cell.borders) return;
        delete cell.borders.top;
        delete cell.borders.right;
        delete cell.borders.bottom;
        delete cell.borders.left;
      };
      for (const cell of this.#selectedTableCells()?.cells ?? []) {
        clearCell(cell);
        if (value === "none") {
          cell.borders = {
            top: { outline: { type: "noFill" } },
            right: { outline: { type: "noFill" } },
            bottom: { outline: { type: "noFill" } },
            left: { outline: { type: "noFill" } },
          };
          continue;
        }
        cell.borders ??= {};
        if (value === "all" || value === "top") cell.borders.top = { width: 12700 };
        if (value === "all" || value === "bottom") cell.borders.bottom = { width: 12700 };
        if (value === "all" || value === "left") cell.borders.left = { width: 12700 };
        if (value === "all" || value === "right") cell.borders.right = { width: 12700 };
      }
    });
  }

  /** The live grid range for Layout commands: an explicit block selection, or
   *  the cell under the text caret. */
  #tableActiveRange(): TableSelectionRange | null {
    if (this.#tableSelection) return this.#tableSelection;
    return this.#tableEdit
      ? { anchor: { ...this.#tableEdit }, head: { ...this.#tableEdit } }
      : null;
  }

  /** Insert a fresh PowerPoint row at the selection's near/far edge. */
  #insertTableRow(position: "above" | "below"): void {
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return;
    const row = Math.min(range.anchor.row, range.head.row);
    const at = position === "above" ? row : Math.max(range.anchor.row, range.head.row) + 1;
    const height = found.table.rows[row]?.height;
    this.#mutateSelectedTable((table) => {
      table.rows.splice(at, 0, {
        height,
        cells: Array.from(
          { length: table.columnWidths?.length ?? 3 },
          () => ({ text: "" }) as TableCellOptions,
        ),
      });
    });
    this.#setTableSelection(null);
  }

  /** Insert a fresh column at the selection edge, cloning the official width. */
  #insertTableColumn(position: "left" | "right"): void {
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return;
    const col = Math.min(range.anchor.col, range.head.col);
    const at = position === "left" ? col : Math.max(range.anchor.col, range.head.col) + 1;
    const width = found.table.columnWidths?.[col];
    this.#mutateSelectedTable((table) => {
      for (const row of table.rows) row.cells.splice(at, 0, { text: "" });
      if (table.columnWidths) table.columnWidths.splice(at, 0, width ?? table.columnWidths[0]);
    });
    this.#setTableSelection(null);
  }

  /** Delete the selected rows/columns; deleting the whole grid deletes the
   *  table object, matching PowerPoint's Delete menu. */
  #deleteTablePart(value?: string): void {
    if (value === "table") return this.#deleteSelected();
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return;
    const fromRow = Math.min(range.anchor.row, range.head.row);
    const toRow = Math.max(range.anchor.row, range.head.row);
    const fromCol = Math.min(range.anchor.col, range.head.col);
    const toCol = Math.max(range.anchor.col, range.head.col);
    if (value === "rows" && fromRow === 0 && toRow >= found.table.rows.length - 1)
      return this.#deleteSelected();
    if (
      value === "columns" &&
      fromCol === 0 &&
      toCol >= (found.table.columnWidths?.length ?? found.table.rows[0]?.cells.length ?? 0) - 1
    )
      return this.#deleteSelected();
    this.#mutateSelectedTable((table) => {
      if (value === "rows") {
        table.rows.splice(fromRow, toRow - fromRow + 1);
        return;
      }
      if (value !== "columns") return;
      const selected = new Set(
        Array.from({ length: toCol - fromCol + 1 }, (_, index) => fromCol + index),
      );
      const origins = new Map(tableGridOf(table).origins.map((origin) => [origin.cell, origin]));
      for (const row of table.rows) {
        row.cells = row.cells.filter((cell) => {
          const origin = origins.get(cell);
          if (!origin) return true;
          const covered = Array.from({ length: origin.spanW }, (_, index) => origin.col + index);
          return !covered.some((col) => selected.has(col));
        });
      }
      if (table.columnWidths)
        table.columnWidths = table.columnWidths.filter((_, col) => !selected.has(col));
    });
    this.#setTableSelection(null);
  }

  /** Merge the selected origins into one rectangular `columnSpan`/`rowSpan`
   *  cell; all absorbed source cells are removed from their rows. */
  #mergeSelectedTableCells(): void {
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return;
    const origins = tableGridOf(found.table).origins;
    const slots = tableSelectionCells(found.member, range)
      .map(({ row, col }) =>
        origins.find(
          (origin) =>
            origin.row <= row &&
            row < origin.row + origin.spanH &&
            origin.col <= col &&
            col < origin.col + origin.spanW,
        ),
      )
      .filter((slot): slot is NonNullable<typeof slot> => Boolean(slot));
    const fromCol = Math.min(...slots.map((slot) => slot.col));
    const fromRow = Math.min(...slots.map((slot) => slot.row));
    const target = slots.find((slot) => slot.row === fromRow && slot.col === fromCol);
    if (slots.length < 2 || !target) return;
    const text = slots
      .filter((slot) => slot.cell !== target.cell)
      .map((slot) => cellTextOf(slot.cell))
      .filter(Boolean)
      .join("\n");
    const removed = new Set(slots.filter((slot) => slot.cell !== target.cell).map((s) => s.cell));
    this.#mutateSelectedTable((table) => {
      for (const row of table.rows) row.cells = row.cells.filter((cell) => !removed.has(cell));
      const toCol = Math.max(...slots.map((slot) => slot.col + slot.spanW)) - 1;
      const toRow = Math.max(...slots.map((slot) => slot.row + slot.spanH)) - 1;
      if (toCol > fromCol) target.cell.columnSpan = toCol - fromCol + 1;
      else delete target.cell.columnSpan;
      if (toRow > fromRow) target.cell.rowSpan = toRow - fromRow + 1;
      else delete target.cell.rowSpan;
      if (text) writeText({ kind: "cell", cell: target.cell }, text);
    });
    this.#setTableSelection({
      anchor: { row: target.row, col: target.col },
      head: { row: target.row, col: target.col },
    });
  }

  /** Split the selected merged origin back to its grid rectangle. Missing
   *  continuation slots become real empty cells; parsed merge placeholders are
   *  replaced rather than duplicated. */
  #splitSelectedTableCell(): void {
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return;
    const at = {
      row: Math.min(range.anchor.row, range.head.row),
      col: Math.min(range.anchor.col, range.head.col),
    };
    const origins = tableGridOf(found.table).origins;
    const before = origins.find(
      (slot) =>
        slot.row <= at.row &&
        at.row < slot.row + slot.spanH &&
        slot.col <= at.col &&
        at.col < slot.col + slot.spanW,
    );
    if (!before || (before.spanW === 1 && before.spanH === 1)) return;
    const oldWidth = before.spanW;
    const oldHeight = before.spanH;
    const blanks = Array.from({ length: oldWidth * oldHeight - 1 }, () => ({ text: "" }));
    this.#mutateSelectedTable((table) => {
      delete before.cell.columnSpan;
      delete before.cell.rowSpan;
      const slots = tableGridOf(table).origins;
      const columns = tableGridOf(table).columns;
      table.rows.forEach((row, rowIndex) => {
        const rowSlots = slots.filter((slot) => slot.row === rowIndex);
        const cells: TableCellOptions[] = [];
        for (let col = 0; col < columns; col += 1) {
          const insideTarget =
            rowIndex >= before.row &&
            rowIndex < before.row + oldHeight &&
            col >= before.col &&
            col < before.col + oldWidth;
          if (insideTarget) {
            if (rowIndex === before.row && col === before.col) cells.push(before.cell);
            else cells.push(blanks[(rowIndex - before.row) * oldWidth + (col - before.col) - 1]!);
            continue;
          }
          const slot = rowSlots.find(
            (candidate) =>
              candidate.row === rowIndex &&
              candidate.col <= col &&
              col < candidate.col + candidate.spanW,
          );
          if (slot && slot.col === col) cells.push(slot.cell);
        }
        row.cells = cells;
      });
    });
    this.#setTableSelection({ anchor: at, head: at });
  }

  /** The selected grid rectangle; an active cell is a one-slot rectangle. */
  #tableGridSelection(): {
    fromRow: number;
    toRow: number;
    fromCol: number;
    toCol: number;
  } | null {
    const found = this.#tableMemberOf();
    const range = this.#tableActiveRange();
    if (!found || !range) return null;
    return {
      fromRow: Math.min(range.anchor.row, range.head.row),
      toRow: Math.max(range.anchor.row, range.head.row),
      fromCol: Math.min(range.anchor.col, range.head.col),
      toCol: Math.max(range.anchor.col, range.head.col),
    };
  }

  /** A typed ribbon measure as EMU; a bare number reads as centimetres. */
  #emuOf(value?: string): number | null {
    const text = value?.trim() ?? "";
    const bare = /^-?\d+(?:\.\d+)?$/.exec(text);
    const cm = bare ? Number(text) : Number.NaN;
    const emu = Number.isFinite(cm) ? Math.round(cm * 360000) : measureEmu(text);
    return emu != null && emu > 0 ? emu : null;
  }

  /** Cell Size boxes write row heights / column widths in EMU; missing width
   *  tracks are materialized from the official table width first. */
  #setTableCellSize(kind: "height" | "width", value?: string): void {
    const emu = this.#emuOf(value);
    const selection = this.#tableGridSelection();
    if (emu == null || !selection) return;
    this.#mutateSelectedTable((table) => {
      if (kind === "height") {
        for (let row = selection.fromRow; row <= selection.toRow; row += 1)
          table.rows[row]!.height = emu;
        return;
      }
      const columns = tableGridOf(table).columns;
      table.columnWidths ??= Array.from({ length: columns }, () =>
        Math.round((measureEmu(table.width) ?? 0) / Math.max(1, columns)),
      );
      for (let col = selection.fromCol; col <= selection.toCol; col += 1)
        table.columnWidths[col] = emu;
    });
  }

  /** PowerPoint's distribute commands equalize only the selected range. */
  #distributeTableCells(kind: "rows" | "columns"): void {
    const found = this.#tableMemberOf();
    const selection = this.#tableGridSelection();
    if (!found || !selection) return;
    this.#mutateSelectedTable((table) => {
      if (kind === "rows") {
        const rows = table.rows
          .slice(selection.fromRow, selection.toRow + 1)
          .map((_, index) => found.member.table.rows[selection.fromRow + index]?.heightPx ?? 0);
        const share = Math.round(
          (rows.reduce((sum, height) => sum + height, 0) * 914400) / 96 / Math.max(1, rows.length),
        );
        for (let row = selection.fromRow; row <= selection.toRow; row += 1)
          table.rows[row]!.height = share;
        return;
      }
      const columns = tableGridOf(table).columns;
      table.columnWidths ??= Array.from({ length: columns }, () =>
        Math.round((measureEmu(table.width) ?? 0) / Math.max(1, columns)),
      );
      const selected = table.columnWidths.slice(selection.fromCol, selection.toCol + 1);
      const share = Math.round(
        selected.reduce((sum: number, width) => sum + (measureEmu(width) ?? 0), 0) /
          Math.max(1, selected.length),
      );
      for (let col = selection.fromCol; col <= selection.toCol; col += 1)
        table.columnWidths![col] = share;
    });
  }

  /** Alignment applies as a range command: paragraph alignment across cells,
   *  official `anchor` for vertical alignment. */
  #setTableCellAlignment(value?: string): void {
    const selected = this.#selectedTableCells();
    if (!selected) return;
    const horizontal = ALIGNMENTS.get(
      value === "left"
        ? "align-left"
        : value === "center"
          ? "align-center"
          : value === "right"
            ? "align-right"
            : "",
    );
    const vertical =
      value === "top" || value === "middle" || value === "bottom"
        ? value === "middle"
          ? "center"
          : value
        : null;
    if (!horizontal && !vertical) return;
    this.#mutateSelectedTable((table) => {
      const origins = new Set(tableGridOf(table).origins.map((origin) => origin.cell));
      for (const cell of selected.cells) {
        if (!origins.has(cell)) continue;
        if (horizontal) setParagraphAlignment(cellParagraphsOf(cell), horizontal);
        else if (vertical) cell.verticalAlign = vertical;
      }
    });
  }

  /** Text direction maps PowerPoint's four table commands onto the cell's
   *  official `vert` token; Horizontal removes the override. */
  #setTableCellTextDirection(value?: string): void {
    if (
      value !== "horizontal" &&
      value !== "vertical" &&
      value !== "vertical270" &&
      value !== "wordArtVertical"
    )
      return;
    const selected = this.#selectedTableCells();
    if (!selected) return;
    this.#mutateSelectedTable((table) => {
      const origins = new Set(tableGridOf(table).origins.map((origin) => origin.cell));
      for (const cell of selected.cells) {
        if (!origins.has(cell)) continue;
        if (value === "horizontal") delete cell.vertical;
        else cell.vertical = value as TextVertical;
      }
    });
  }

  /** PowerPoint's preset margin table in EMU; Normal removes per-cell values
   *  so the DrawingML defaults stay authoritative. */
  #setTableCellMargins(value?: string): void {
    if (value !== "normal" && value !== "none" && value !== "narrow" && value !== "wide") return;
    const selected = this.#selectedTableCells();
    if (!selected) return;
    const margins =
      value === "none"
        ? { top: 0, right: 0, bottom: 0, left: 0 }
        : value === "narrow"
          ? { top: 45720, right: 45720, bottom: 45720, left: 45720 }
          : { top: 91440, right: 182880, bottom: 91440, left: 182880 };
    this.#mutateSelectedTable((table) => {
      const origins = new Set(tableGridOf(table).origins.map((origin) => origin.cell));
      for (const cell of selected.cells) {
        if (!origins.has(cell)) continue;
        if (value === "normal") delete cell.margins;
        else cell.margins = { ...margins };
      }
    });
  }

  #pickPicture(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#picture-input")?.click();
  }

  #pickMedia(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#media-input")?.click();
  }

  #pickObject(): void {
    if (!this.#pres || this.#objectDialog?.open) return;
    const dialog = document.createElement("dialog");
    dialog.className = "insert-dialog";
    dialog.innerHTML = `
      <div class="dialog-head"><strong>${escapeHtml(t("ppt.object.title", this))}</strong><button data-dialog-close>×</button></div>
      <div class="dialog-body">
        <fieldset class="object-source">
          <legend>${escapeHtml(t("ppt.object.source", this))}</legend>
          <label><input type="radio" name="object-source" value="embed" checked> ${escapeHtml(t("ppt.object.embed", this))}</label>
          <label><input type="radio" name="object-source" value="link"> ${escapeHtml(t("ppt.object.link", this))}</label>
        </fieldset>
        <label class="dialog-field" data-object-mode="embed"><span>${escapeHtml(t("ppt.object.file", this))}</span><input id="object-file-name" readonly placeholder="${escapeHtml(t("ppt.object.choose", this))}"></label>
        <label class="dialog-field" data-object-mode="link" hidden><span>${escapeHtml(t("ppt.object.url", this))}</span><input id="object-url" placeholder="https://example.com/report.xlsx"></label>
        <label class="hf-option"><input id="object-icon" type="checkbox" checked> ${escapeHtml(t("ppt.object.show-as-icon", this))}</label>
        <label class="hf-option" data-object-auto hidden><input id="object-auto" type="checkbox"> ${escapeHtml(t("ppt.object.auto-update", this))}</label>
      </div>
      <div class="dialog-actions"><button data-dialog-cancel>${escapeHtml(t("ppt.dialog.cancel", this))}</button><button data-object-pick>${escapeHtml(t("ppt.dialog.insert", this))}</button></div>
    `;
    dialog.addEventListener("change", (event) => {
      const target = event.target as HTMLInputElement;
      if (target.name === "object-source") {
        const link = target.value === "link";
        this.#objectLink = link;
        for (const item of dialog.querySelectorAll<HTMLElement>("[data-object-mode]"))
          item.hidden = (item.dataset.objectMode === "link") !== link;
        dialog.querySelector<HTMLElement>("[data-object-auto]")!.hidden = !link;
      } else if (target.id === "object-file") {
        const file = target.files?.[0];
        if (dialog.querySelector<HTMLInputElement>("#object-file-name") && file)
          dialog.querySelector<HTMLInputElement>("#object-file-name")!.value = file.name;
      }
    });
    dialog.addEventListener("click", async (event) => {
      const target = event.target as HTMLElement;
      if (target.closest("[data-dialog-close]") || target.closest("[data-dialog-cancel]"))
        return dialog.close();
      if (!target.closest("[data-object-pick]")) return;
      const showAsIcon = dialog.querySelector<HTMLInputElement>("#object-icon")?.checked === true;
      this.#objectShowAsIcon = showAsIcon;
      this.#objectAutoUpdate =
        dialog.querySelector<HTMLInputElement>("#object-auto")?.checked === true;
      if (this.#objectLink) {
        const url = dialog.querySelector<HTMLInputElement>("#object-url")?.value.trim();
        if (!url) return;
        const name = decodeURIComponent(url.split(/[?#]/)[0]!.split("/").pop() || "Object");
        this.#insertChild(
          makeObject(
            this.#pres!.widthPx,
            this.#pres!.heightPx,
            new Uint8Array(),
            await objectIconDataUrl(name),
            name,
            objectProgIdOf(name),
            { url, autoUpdate: this.#objectAutoUpdate },
            showAsIcon,
          ),
        );
        dialog.close();
        return;
      }
      this.shadowRoot?.querySelector<HTMLInputElement>("#object-input")?.click();
    });
    this.#showModalDialog(dialog, () => (this.#objectDialog = null));
    this.#objectDialog = dialog;
  }

  /** File-input object path: native bytes enter the real OLE embed; the
   *  generated icon is the `p:pic` PowerPoint needs to reopen the frame. */
  readonly #onObjectChange = async (event: Event): Promise<void> => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    const pres = this.#pres;
    if (!file || !pres) return;
    const icon = await objectIconDataUrl(file.name);
    this.#insertChild(
      makeObject(
        pres.widthPx,
        pres.heightPx,
        new Uint8Array(await file.arrayBuffer()),
        icon,
        file.name,
        objectProgIdOf(file.name),
        undefined,
        this.#objectShowAsIcon,
      ),
    );
    this.#objectDialog?.close();
  };

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
    // A live textarea's selection addresses styled ranges; capture it before
    // the write-back/exit because focus changes can move the native selection.
    const editor = this.#textEditor;
    const selection =
      editor == null
        ? null
        : {
            start: editor.selectionStart ?? 0,
            end: editor.selectionEnd ?? editor.selectionStart ?? 0,
          };
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
    let applied = false;
    if (selection) {
      const source: TextEditSource = cell
        ? { kind: "cell", cell }
        : { kind: "shape", child: (child as Extract<SlideChild, { shape: unknown }>).shape };
      applied = formatText(source, selection.start, selection.end, name, value);
    } else {
      switch (name) {
        case "bold":
          toggleRunFlag(paragraphs, "bold");
          applied = true;
          break;
        case "italic":
          toggleRunFlag(paragraphs, "italic");
          applied = true;
          break;
        case "underline":
          toggleRunStyle(paragraphs, "underline", "single");
          applied = true;
          break;
        case "strike":
          toggleRunStyle(paragraphs, "strike", "singleStrike");
          applied = true;
          break;
        case "font-face":
          if (value) setRunFont(paragraphs, value);
          applied = Boolean(value);
          break;
        case "font-size": {
          const size = Number(value);
          if (Number.isFinite(size) && size > 0) setRunSize(paragraphs, size);
          applied = Number.isFinite(size) && size > 0;
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
          applied = Boolean(
            alignment ||
            name === "list" ||
            name === "numbering" ||
            LINE_SPACING_OF.has(value ?? ""),
          );
        }
      }
    }
    if (!applied) return;
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

  /** Stamp Cell Size from the selected grid range; mixed extents stay blank. */
  #syncTableCellSizeControls(): void {
    const root = this.shadowRoot;
    const found = this.#tableMemberOf();
    const selection = this.#tableGridSelection();
    if (!root || !found || !selection) return;
    const commonEmu = (values: unknown[]): number | null => {
      const known = values
        .map((value) => measureEmu(value))
        .filter((value): value is number => value != null);
      return known.length === values.length &&
        known.length > 0 &&
        known.every((value) => value === known[0])
        ? known[0]
        : null;
    };
    const height = commonEmu(
      Array.from({ length: selection.toRow - selection.fromRow + 1 }, (_, index) => {
        const row = selection.fromRow + index;
        return (
          found.table.rows[row]?.height ??
          Math.round((found.member.table.rows[row]?.heightPx ?? 0) * (914400 / 96))
        );
      }),
    );
    const columns = tableGridOf(found.table).columns;
    const width = commonEmu(
      Array.from({ length: selection.toCol - selection.fromCol + 1 }, (_, index) => {
        const column = selection.fromCol + index;
        return (
          found.table.columnWidths?.[column] ??
          (Math.round((found.member.table.columnWidthsPx[column] ?? 0) * (914400 / 96)) ||
            Math.round((measureEmu(found.table.width) ?? 0) / Math.max(1, columns)))
        );
      }),
    );
    const format = (emu: number | null): string =>
      emu == null ? "" : `${(emu / 360000).toFixed(2)} cm`;
    for (const [event, value] of [
      ["cell-height", format(height)],
      ["cell-width", format(width)],
    ] as const) {
      const box = root.querySelector<HTMLElement>(`docen-ribbon-input[event='${event}']`);
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
      if (event.button === 0) this.#advanceShow();
      return;
    }
    if (this.#shapeDrawer?.armed && this.#shapeDrawer.startFromPointer(event)) return;
    if (this.#drawTool === "eraser" && event.button === 0) return this.#startEraser(event);
    if (this.#drawTool === "pen" && event.button === 0) return this.#startPen(event);
    const presJson = this.#presJson;
    if (!presJson || event.button !== 0) return;
    const point = this.#stagePointOf(event);
    if (!point) return;
    // The selected table's grips outrank a live cell edit: DOCX can ask for a
    // row/column/whole-table selection without first leaving the bridge.
    const selectedTable = this.#tableMemberOf();
    const grip =
      selectedTable?.slide === point.slide
        ? tableGripAt(selectedTable.member, point.x, point.y)
        : null;
    if (grip?.clickable) {
      event.preventDefault();
      event.stopPropagation();
      this.#selectTableGrip(grip.kind, grip.index);
      return;
    }
    const child = hitSlide(slideHits(presJson.slides?.[point.slide] ?? {}), point.x, point.y);
    // A press outside the object under edit leaves the edit (Word's rule);
    // one inside just moves the caret — the textarea keeps the session. In a
    // table edit a press on another cell of the same table hops there.
    if (this.#textEditor) {
      if (this.#selection?.slide !== point.slide || this.#selection.child !== child)
        this.#select(null);
      else if (this.#tableEdit) {
        // The editor owns cross-cell drags, but a press in the same cell
        // keeps the browser's native caret and text selection in charge.
        const table = this.#tableMemberOf();
        const rect = table ? this.#cellRectAt(table.member, point.x, point.y) : null;
        if (table && rect) {
          const edit = this.#tableEdit!;
          const sameCell = edit.row === rect.cell.row && edit.col === rect.cell.col;
          const activeEditor = this.#textEditor;
          this.#startTableCellDrag(event, table, () => {
            if (!sameCell) {
              this.#moveTableCellEditing(point.x, point.y, undefined, {
                caret: { clientX: event.clientX, clientY: event.clientY },
              });
              return;
            }
            this.#setTableSelection(null);
            activeEditor?.focus();
            if (activeEditor) this.#placeTableCellCaret(activeEditor, event.clientX, event.clientY);
          });
        }
      }
      return;
    }
    if (child < 0) return this.#select(null);
    const target = presJson.slides?.[point.slide]?.children?.[child];
    // A table body opens Word's cell edit on the first press; its grips are
    // the only way to ask for a row/column/whole-table selection. Nested
    // tables descend into their member before resolving the projected grid.
    const nested = target && "group" in target ? memberAt(target, point.x, point.y) : null;
    const nestedLeaf = nested && this.#childAt(target!, nested.path);
    const tableLeaf = target && "table" in target ? target : nestedLeaf;
    if (tableLeaf && "table" in tableLeaf) {
      this.#select({
        slide: point.slide,
        child,
        ...(target && "group" in target && nested ? { member: nested.path } : {}),
      });
      const table = this.#tableMemberOf();
      const rect = table ? this.#cellRectAt(table.member, point.x, point.y) : null;
      if (!table || !rect) return;
      event.preventDefault();
      // With a live block selection, the press is Word's extend-selection
      // gesture; only a release without movement collapses to a caret.
      if (this.#tableSelection) {
        this.#startTableCellDrag(event, table, () =>
          this.#enterTableCellEditing(point.x, point.y, undefined, {
            clientX: event.clientX,
            clientY: event.clientY,
          }),
        );
        return;
      }
      this.#enterTableCellEditing(point.x, point.y, undefined, {
        clientX: event.clientX,
        clientY: event.clientY,
      });
      this.#startTableCellDrag(event, table);
      return;
    }
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
        return this.#enterTableCellEditing(point.x, point.y, undefined, {
          clientX: event.clientX,
          clientY: event.clientY,
        });
      // A group's table member opens its cell edit — the first click of the
      // double-click already descended into the member.
      if (child >= 0 && target && "group" in target) {
        const hit = memberAt(target, point.x, point.y);
        const leaf = hit && this.#childAt(target, hit.path);
        if (hit && leaf && "table" in leaf)
          return this.#enterTableCellEditing(point.x, point.y, undefined, {
            clientX: event.clientX,
            clientY: event.clientY,
          });
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
        this.#advanceShow();
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
    if (event.key === "Escape" && this.#drawTool) return this.#armDrawTool("select");
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
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      return void this.#saveAs();
    }
    if ((event.key === "Delete" || event.key === "Backspace") && this.#selection) {
      event.preventDefault();
      if (this.#tableSelection) return this.#clearSelectedTableCells();
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
