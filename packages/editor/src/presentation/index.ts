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
  generatePresentation,
  parsePresentation,
  projectPresentation,
  type PresentationOptions,
  type ProjectedPresentation,
  type SlideChild,
} from "@docen/pptx";
import { customElement, observable } from "@microsoft/fast-element";
import type { DataType } from "@office-open/core";
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
import { escapeHtml, presentationStyles, presentationTemplate } from "./chrome";
import {
  firstRunSizeOf,
  insertSlideAt,
  makePicture,
  makeTextBox,
  setParagraphAlignment,
  setRunFont,
  setRunSize,
  shapeTextOf,
  toggleRunFlag,
  toggleRunStyle,
  writeShapeText,
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

/** Undo steps kept before the oldest falls off. */
const EDIT_LIMIT = 50;

/** Text/paragraph commands over the selected shape (font/size carry the
 *  combobox value through detail.value). */
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
]);

/** Ribbon alignment command → TextAlignment value. */
const ALIGNMENTS: ReadonlyMap<string, "left" | "center" | "right" | "justify"> = new Map([
  ["align-left", "left"],
  ["align-center", "center"],
  ["align-right", "right"],
  ["justify", "justify"],
]);

/** Browser MIME → the JSON picture type (pptx's supported raster set). */
const PICTURE_TYPES: ReadonlyMap<string, "png" | "jpg" | "gif" | "bmp"> = new Map([
  ["image/png", "png"],
  ["image/jpeg", "jpg"],
  ["image/gif", "gif"],
  ["image/bmp", "bmp"],
]);

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

/** Commands with a live handler — everything else greys out (the honest
 *  ribbon). Add-in commands join this set at render time. */
const WIRED_COMMANDS: ReadonlySet<string> = new Set([
  "zoom-in",
  "zoom-out",
  "zoom-100",
  "undo",
  "redo",
  "save-as",
  "new-slide",
  "text-box",
  "insert-picture",
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
  #selection: { slide: number; child: number } | null = null;
  /** The in-place shape-text editor: the textarea floats over the shape while
   *  it holds the session (its position; the text is read back on exit). */
  #textEditor: HTMLTextAreaElement | null = null;
  #edits: DeckEdit[] = [];
  #editIndex = -1;
  #langObserver?: MutationObserver;
  #unsubLang?: () => void;
  #scrollRaf?: number;

  /** Template refs into the thumbnails panel (see chrome.ts). */
  @observable thumbStrip?: HTMLElement;
  @observable thumbSelection?: HTMLElement;

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
    this.#area()?.addEventListener("scroll", this.#onScroll);
    this.thumbStrip?.addEventListener("click", this.#onThumbClick);
    this.#stage.addEventListener("pointerdown", this.#onStagePointerDown);
    this.#stage.addEventListener("dblclick", this.#onStageDblClick);
    document.addEventListener("keydown", this.#onKeyDown);
    // The selection frame lives over the slide surface (the canvas element
    // Leafer mounts is recreated per deck render — the overlay survives it).
    this.#overlay = new DrawingOverlay({
      scale: () => this.#zoom / 100,
      applyBox: (box) => this.#applyGesture((child) => resizeChild(child, this.#slideBoxOf(box))),
      applyOffset: (dx, dy) => this.#applyGesture((child) => offsetChild(child, dx, dy)),
      applyRotation: (delta) => this.#applyGesture((child) => rotateChild(child, delta)),
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

  /** Title-bar menu items carry their action in `data-event`. */
  readonly #onChange = (event: Event): void => {
    const name = (event.target as HTMLElement)?.dataset?.event;
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
    else if (name === "text-box") this.#insertTextBox();
    else if (name === "insert-picture") this.#pickPicture();
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
   *  box plus the slide's offset in the strip) plus its spin. */
  #selectedBox(): { box: Box; rotation: number } | null {
    const sel = this.#selection;
    const presJson = this.#presJson;
    if (!sel || !presJson) return null;
    const slide = presJson.slides?.[sel.slide];
    const hit = slideHits(slide ?? {}).find((h) => h.child === sel.child);
    if (!hit) return null;
    const pres = this.#pres!;
    return {
      box: {
        x: hit.box.x,
        y: hit.box.y + sel.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX,
        width: hit.box.width,
        height: hit.box.height,
      },
      rotation: hit.rotation,
    };
  }

  /** Select a slide object (or clear). The frame shows the strip-space box. */
  #select(sel: { slide: number; child: number } | null): void {
    if (this.#textEditor) this.#exitTextEditing(true);
    this.#selection = sel;
    if (!sel) return this.#overlay?.hide();
    const hit = this.#selectedBox();
    if (hit) this.#overlay?.show(hit.box, hit.rotation);
    else this.#overlay?.hide();
  }

  /** Double-click on a shape: float a textarea over its text area and hand
   *  the session to it (PowerPoint's in-place text edit). The overlay's
   *  resize handles step aside for the duration. */
  #enterTextEditing(): void {
    const sel = this.#selection;
    const pres = this.#pres;
    const presJson = this.#presJson;
    const child = presJson?.slides?.[sel?.slide ?? -1]?.children?.[sel?.child ?? -1];
    const text = child ? shapeTextOf(child) : null;
    if (!sel || !pres || !presJson || !child || text === null) return;
    const hit = slideHits(presJson.slides?.[sel.slide] ?? {}).find((h) => h.child === sel.child);
    if (!hit) return;
    const scale = this.#zoom / 100;
    const stripY = hit.box.y + sel.slide * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
    const editor = document.createElement("textarea");
    editor.className = "shape-text-editor";
    editor.value = text;
    Object.assign(editor.style, {
      left: `${hit.box.x * scale}px`,
      top: `${stripY * scale}px`,
      width: `${hit.box.width * scale}px`,
      height: `${hit.box.height * scale}px`,
      fontSize: `${((firstRunSizeOf(child) * 4) / 3) * scale}px`,
    });
    editor.addEventListener("keydown", (event) => {
      // Escape leaves the edit (committing); typing keys stay in the textarea.
      if (event.key === "Escape") {
        event.stopPropagation();
        this.#exitTextEditing(true);
      }
      event.stopPropagation();
    });
    editor.addEventListener("input", () => {
      // Grow with the text so multi-line entries stay visible (PowerPoint's
      // box autofit — height only, the width is the shape's).
      editor.style.height = "auto";
      editor.style.height = `${editor.scrollHeight}px`;
    });
    this.#canvasHost().append(editor);
    editor.focus();
    editor.select();
    this.#textEditor = editor;
    this.#overlay?.hide();
  }

  /** Leave the text edit. `write` commits the textarea's text back into the
   *  shape (an undo-able edit when it changed); plain exits just drop it. */
  #exitTextEditing(write: boolean): void {
    const editor = this.#textEditor;
    this.#textEditor = null;
    editor?.remove();
    const sel = this.#selection;
    const child = this.#presJson?.slides?.[sel?.slide ?? -1]?.children?.[sel?.child ?? -1];
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

  /** Commit a drag gesture to the source child, record it, and repaint. */
  #applyGesture(edit: (child: SlideChild) => void): void {
    const sel = this.#selection;
    const slide = this.#presJson?.slides?.[sel?.slide ?? -1];
    const child = slide?.children?.[sel!.child];
    if (!child) return;
    const before = captureGeometry(child);
    edit(child);
    const after = captureGeometry(child);
    this.#pushEdit({
      undo: () => {
        restoreGeometry(child, before);
        this.#reproject(sel!.slide);
      },
      redo: () => {
        restoreGeometry(child, after);
        this.#reproject(sel!.slide);
      },
    });
    this.#reproject(sel!.slide);
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

  #pickPicture(): void {
    this.shadowRoot?.querySelector<HTMLInputElement>("#picture-input")?.click();
  }

  /** File-input picture path: read the file as a data URL, measure its
   *  natural size, and drop a centered picture child on the active slide. */
  readonly #onPictureChange = async (event: Event): Promise<void> => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    const pres = this.#pres;
    const type = file ? PICTURE_TYPES.get(file.type) : undefined;
    if (!file || !pres || !type) return;
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
  };

  /** Apply a font/paragraph command to the selected shape's text body. The
   *  undo pair clones the whole body (geometry snapshots don't reach text);
   *  a command that changed nothing records nothing. */
  #applyTextFormat(name: string, value?: string): void {
    // Format lands on the committed text: an in-flight edit writes back
    // first, or the exit-time write-back would clobber the format.
    this.#exitTextEditing(true);
    const sel = this.#selection;
    const child = this.#presJson?.slides?.[sel?.slide ?? -1]?.children?.[sel?.child ?? -1];
    if (!child || !("shape" in child)) return;
    const shape = child.shape;
    const body = shape.textBody;
    if (!body) return;
    const before = structuredClone(body);
    switch (name) {
      case "bold":
        toggleRunFlag(body, "bold");
        break;
      case "italic":
        toggleRunFlag(body, "italic");
        break;
      case "underline":
        toggleRunStyle(body, "underline", "single");
        break;
      case "strike":
        toggleRunStyle(body, "strike", "singleStrike");
        break;
      case "font-face":
        if (value) setRunFont(body, value);
        break;
      case "font-size": {
        const size = Number(value);
        if (Number.isFinite(size) && size > 0) setRunSize(body, size);
        break;
      }
      default: {
        const alignment = ALIGNMENTS.get(name);
        if (alignment) setParagraphAlignment(body, alignment);
      }
    }
    if (JSON.stringify(body) === JSON.stringify(before)) return;
    const after = structuredClone(body);
    this.#pushEdit({
      undo: () => {
        shape.textBody = structuredClone(before);
        this.#reproject(sel!.slide);
      },
      redo: () => {
        shape.textBody = structuredClone(after);
        this.#reproject(sel!.slide);
      },
    });
    this.#reproject(sel!.slide);
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

  readonly #onStagePointerDown = (event: PointerEvent): void => {
    const pres = this.#pres;
    const presJson = this.#presJson;
    if (!pres || !presJson || event.button !== 0) return;
    const rect = this.#stage.getBoundingClientRect();
    if (rect.width === 0) return;
    const x = ((event.clientX - rect.left) / rect.width) * pres.widthPx;
    const stripY =
      ((event.clientY - rect.top) / rect.height) *
      (pres.slides.length * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX);
    const slide = Math.floor(stripY / (pres.heightPx + SLIDE_GAP_PX));
    if (slide < 0 || slide >= pres.slides.length) return;
    const localY = stripY - slide * (pres.heightPx + SLIDE_GAP_PX) - SLIDE_GAP_PX;
    const child = hitSlide(slideHits(presJson.slides?.[slide] ?? {}), x, localY);
    // A press outside the shape under edit leaves the edit (Word's rule);
    // one inside just moves the caret — the textarea keeps the session.
    if (this.#textEditor) {
      if (this.#selection?.slide !== slide || this.#selection.child !== child) this.#select(null);
      return;
    }
    if (child < 0) return this.#select(null);
    if (this.#selection?.slide === slide && this.#selection.child === child)
      return this.#overlay?.beginMove(event.clientX, event.clientY);
    this.#select({ slide, child });
  };

  readonly #onStageDblClick = (): void => {
    if (this.#textEditor) return;
    this.#enterTextEditing();
  };

  readonly #onKeyDown = (event: KeyboardEvent): void => {
    // Inside the text edit the textarea owns every key: Enter/Delete type,
    // its Escape handler commits and leaves; undo is the browser's (the
    // textarea's own typing history).
    if (this.#textEditor) return;
    if (event.key === "Escape" && this.#selection) return this.#select(null);
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
    if ((event.key === "Delete" || event.key === "Backspace") && this.#selection) {
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
