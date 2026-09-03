import {
  paintColumnSeparators,
  paintFootnotes,
  paintFurnitureStack,
  paintGridlines,
  paintLineNumbers,
  paintScene,
  releasePinnedImages,
  type DrawingHitBox,
  type LineNumberMark,
  type PaintContext,
} from "@docen/core";
import type {
  LayoutBlock,
  ProjectedColumns,
  ProjectedFlowBox,
  ProjectedLineNumbers,
  ProjectedPageBackground,
  ProjectedPageBorder,
  ProjectedPageBorders,
  ProjectedPageFurniture,
} from "@docen/layout";
/**
 * Canvas stage — the page layer of the canvas document component.
 *
 * One LeaferJS App per page slot, created/destroyed by an IntersectionObserver
 * as pages approach the viewport (scrolling stays native DOM one level up —
 * the pages never consume the wheel). Repainting a page reuses its App: the
 * tree is cleared and rebuilt, keeping canvas creation (the expensive part)
 * off the per-sync path. App, not the bare Leafer class: `editor` in the
 * config triggers the tree+sky layer creation (without it no canvas appears).
 *
 * Demo-isolated for now: this module mounts into a plain container the demo
 * owns and injects its own page styles; it wires into <docen-document> only
 * after the editing milestones (M-R2+) land.
 */
import type { FlowPage, FontMetrics, LaidOutStackItem } from "@docen/layout";
import { stackBlocks, TextMeasurer } from "@docen/layout";
import { App, Debug, Group, Line, Rect, Text, type IGroup } from "leafer-ui";

import { computeLineNumbers } from "./line-numbers";

const PAGE_GAP = 24;

/** One laid furniture slot: the paint-ready stack and its laid height. */
export interface LaidFurnitureSlot {
  stack: readonly LaidOutStackItem[];
  heightPx: number;
}

/** A page's persistent paint layers, in z-order. The two furniture groups
 *  survive body-only repaints — headers/footers are per-section constants in
 *  Word, and re-creating their image leaves on every keystroke re-decode gap
 *  that reads as header flicker while deleting. */
interface PageLayers {
  /** Body behind-text floats (under furniture). */
  behind: Group;
  /** Header/footer behind-text drawings (watermarks). */
  furnitureBehind: Group;
  /** Header/footer stories. */
  furnitureBody: Group;
  /** Body content plus its deferred in-front floats. */
  body: Group;
}

/** One page slot: its DOM frame, its Leafer App, and the layers its last
 *  full paint created (null while a story edit paints flat, after an App
 *  recycle, or before the first paint). */
interface PageSlot {
  el: HTMLElement;
  app: App | null;
  layers: PageLayers | null;
}

/** One section's laid furniture slots — [default, first, even]. */
export interface LaidFurnitureSection {
  header: (LaidFurnitureSlot | undefined)[];
  footer: (LaidFurnitureSlot | undefined)[];
}

/** One section's paint inputs: the page geometry its pages paginate against
 *  and the headers/footers its pages display. */
export interface CanvasStageSection {
  flow: ProjectedFlowBox;
  /** The section's page borders (w:pgBorders), absent when none. */
  pageBorders?: ProjectedPageBorders;
  /** The section's line numbering (w:lnNumType), absent when none. */
  lineNumbers?: ProjectedLineNumbers;
  /** The section's columns (w:cols), absent for a single-column section. */
  columns?: ProjectedColumns;
  /** Headers/footers for this section's pages (absent = none). */
  furniture?: ProjectedPageFurniture;
  /** The slots of `furniture` laid out once (layFurnitureSections) — the
   *  page insets push the body by these heights and the painter draws these
   *  same stacks, so push-down == painted band height by construction. */
  furnitureLaid?: LaidFurnitureSection;
  /** Footnote definitions (absent when document has no footnotes). */
  footnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote definitions (absent when document has no endnotes). */
  endnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
}

/** [default, first, even] slot pick order. */
const FURNITURE_SLOTS = [0, 1, 2] as const;

/** Lay every furniture slot once, at its section's content width — the
 *  single pass both consumers share. No grid context: Word keeps
 *  header/footer paragraphs at natural line height (the body docGrid does
 *  not apply to the furniture story). */
export function layFurnitureSections(
  sections: readonly CanvasStageSection[],
  metrics: FontMetrics,
): (LaidFurnitureSection | undefined)[] {
  const measurer = new TextMeasurer(metrics);
  return sections.map((section) => {
    const f = section.furniture;
    if (!f) return undefined;
    const lay = (blocks: ProjectedPageFurniture["header"]): LaidFurnitureSlot | undefined => {
      if (!blocks) return undefined;
      const laid = stackBlocks(blocks, section.flow.contentWidthPx, undefined, measurer);
      return { stack: laid.stack, heightPx: laid.heightPx };
    };
    return {
      header: FURNITURE_SLOTS.map((slot) => lay([f.header, f.firstHeader, f.evenHeader][slot])),
      footer: FURNITURE_SLOTS.map((slot) => lay([f.footer, f.firstFooter, f.evenFooter][slot])),
    };
  });
}

/** Font metrics (baseline ratios for half-leading) + per-section geometry. */
export interface CanvasStageContext {
  metrics: FontMetrics;
  /** One entry per document section — each page paints with the box and
   *  furniture of the section {@link CanvasStageContext.sectionOfPage} maps
   *  it to. */
  sections: CanvasStageSection[];
  /** Page index → index into `sections` (parallel to the page list). */
  sectionOfPage: number[];
  /** Page background (w:background — base color + optional pattern tile). */
  background?: ProjectedPageBackground;
  /** The break rows' labels in the UI language (Word paints them localized). */
  marksLabels?: { pageBreak?: string; sectionBreak?: string };
}

export class CanvasStage {
  readonly shell: HTMLElement;
  private readonly slots: PageSlot[] = [];
  private readonly io: IntersectionObserver;
  private pages: FlowPage[] = [];
  /** Per-page line-number labels (sync's one-pass count; empty pages of
   *  unnumbered sections carry an empty list). */
  private lineNumberMarks = new Map<number, LineNumberMark[]>();

  /** The section a page belongs to (its flow box + furniture). */
  private sectionAt(page: number): CanvasStageSection {
    const i = this.ctx.sectionOfPage[page] ?? 0;
    return this.ctx.sections[i] ?? this.ctx.sections[0]!;
  }

  constructor(
    stage: HTMLElement,
    private readonly ctx: CanvasStageContext,
  ) {
    // The page layer: a fit-content centered stack inside the host's scroll
    // container. NOT a scroll surface itself.
    this.shell = document.createElement("div");
    this.shell.className = "canvas-pages";
    Object.assign(this.shell.style, {
      width: "fit-content",
      margin: "0 auto",
      padding: "32px 0",
    } satisfies Partial<CSSStyleDeclaration>);
    // LeaferJS's interaction layer preventDefaults wheel at its view
    // (app-config flags do not stop that), which would kill the browser's
    // default scrolling of the outer container. Cut the event before it
    // reaches the canvases — the pages scroll as one document surface.
    this.shell.addEventListener("wheel", (event) => event.stopPropagation(), { capture: true });
    stage.append(this.shell);
    this.io = new IntersectionObserver(
      (records) => {
        for (const record of records) {
          const slot = this.slots.find((s) => s.el === record.target);
          if (!slot) continue;
          if (record.isIntersecting) this.ensure(slot);
          else if (slot.app) {
            slot.app.destroy();
            slot.app = null;
            // The layers belonged to the destroyed tree — drop the stale
            // reference so the next paint rebuilds rather than addressing
            // groups of a dead App.
            slot.layers = null;
          }
        }
      },
      // The IO root is the SCROLL CONTAINER (the stage host's parent), not the
      // viewport: rootMargin only widens the root's clip rect, and an
      // intermediate overflow:auto box re-clips it — pages below the scroller's
      // visible area would never pre-render against a viewport root. The
      // margins (order top/right/bottom/left) pre-render ahead of BOTH scroll
      // directions: downward for fast scrolls, and upward because deletion
      // pulls content up — the viewport effectively scans upward through the
      // doc, and a one-sided margin there shows blank pages while holding
      // Backspace.
      { root: stage.parentElement, rootMargin: "150% 0px 150% 0px" },
    );
    // Leafer samples devicePixelRatio at App creation and misses a cross-
    // monitor / browser-zoom change, leaving the canvas CSS-stretched (blurry
    // text and bitmaps). Watch the ratio and re-render every live App at the
    // new one — rare and cheap.
    this.watchPixelRatio();
  }

  private dprMedia: MediaQueryList | null = null;
  /** Current zoom level (percent). The zoom IS the layout: slots are sized to
   *  the scaled page directly (no CSS zoom anywhere) so each canvas bitmaps
   *  at exactly its on-screen pixel size. */
  private zoomPercent = 100;

  private readonly dprChange = () => {
    this.watchPixelRatio();
    this.applyZoom();
  };

  private watchPixelRatio(): void {
    this.dprMedia?.removeEventListener("change", this.dprChange);
    this.dprMedia = matchMedia(`(resolution: ${devicePixelRatio}dppx)`);
    this.dprMedia.addEventListener("change", this.dprChange);
  }

  private get factor(): number {
    return this.zoomPercent / 100;
  }

  /** The zoom scale as the edit overlays need it (semantic px → screen px). */
  scale(): number {
    return this.factor;
  }

  /** The current zoom level (percent). */
  get zoom(): number {
    return this.zoomPercent;
  }

  /** A page's on-screen CSS size at the current zoom, snapped so the bitmap
   *  (css × pixelRatio) covers whole device pixels — the canvas then composites
   *  1:1 with no resampling, which is what keeps text sharp at EVERY zoom
   *  level (fractional scales like 110% otherwise resample and blur). */
  private pageCss(px: number): number {
    return Math.round(px * this.factor * devicePixelRatio) / devicePixelRatio;
  }

  /** Bitmap-width cap: A4 at 500% on a 2× screen would be a ~360MB bitmap per
   *  live page. Past the cap the canvas re-upsamples (soft) instead of
   *  exhausting memory. */
  private static readonly BITMAP_CAP = 3600;

  private renderPixelRatio(flow: ProjectedFlowBox): number {
    const w = flow.pageWidthPx * this.factor;
    return Math.min(devicePixelRatio, CanvasStage.BITMAP_CAP / Math.max(w, 1));
  }

  /** Zoom change → resize every slot to the scaled page and re-render its
   *  canvas at the matching resolution. `app.resize` re-uses the App (no
   *  destroy/create churn while dragging the zoom slider); `tree.scale`
   *  keeps all paint coordinates in unzoomed page px. */
  setZoom(pct: number): void {
    if (pct === this.zoomPercent) return;
    this.zoomPercent = pct;
    this.applyZoom();
  }

  /** Formatting marks (Word's ¶ toggle) — a paint-time flag: flipping it
   *  repaints every live page with marks drawn (or dropped). */
  #showMarks = false;

  setShowMarks(on: boolean): void {
    if (on === this.#showMarks) return;
    this.#showMarks = on;
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
    }
  }

  /** Document-grid overlay (Word's View → Gridlines) — a paint-time flag like
   *  the formatting marks: flipping it repaints every live page. */
  #showGridlines = false;

  get showGridlines(): boolean {
    return this.#showGridlines;
  }

  setShowGridlines(on: boolean): void {
    if (on === this.#showGridlines) return;
    this.#showGridlines = on;
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
    }
  }

  /** Leafer's engine debug flags (element wireframes / hit areas / repaint
   *  regions + engine logs) — global statics on leafer-ui, so one call covers
   *  every live page App. `mode`: "bounds" | "hit" | "repaint" | "on";
   *  anything else clears. */
  setDebug(mode: string | null | undefined): void {
    const on = mode === "on" || mode === "bounds" || mode === "hit" || mode === "repaint";
    Debug.enable = on;
    Debug.showBounds = mode === "bounds" ? true : mode === "hit" ? "hit" : false;
    Debug.showRepaint = mode === "repaint";
  }

  /** Ruler visibility (Word's View → Ruler). The rulers are frame-level DOM
   *  overlays (see {@link applyRulers}), so this just mounts/unmounts them. */
  #showRuler = false;

  get showRuler(): boolean {
    return this.#showRuler;
  }

  setShowRuler(on: boolean): void {
    if (on === this.#showRuler) return;
    this.#showRuler = on;
    for (const [index, slot] of this.slots.entries()) {
      const frame = slot.el.parentElement;
      if (frame) this.applyRulers(frame, index);
    }
  }

  /** The break rows' labels (locale change): repainting is deferred to the
   *  next marks-visible repaint when marks are off, immediate when on. */
  setMarksLabels(labels: { pageBreak: string; sectionBreak: string }): void {
    const cur = this.ctx.marksLabels;
    if (cur?.pageBreak === labels.pageBreak && cur?.sectionBreak === labels.sectionBreak) return;
    this.ctx.marksLabels = labels;
    if (!this.#showMarks) return;
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
    }
  }

  private applyZoom(): void {
    // Sections can carry different paper sizes — each slot sizes to its own.
    for (const [index, slot] of this.slots.entries()) {
      const flow = this.sectionAt(index).flow;
      const w = this.pageCss(flow.pageWidthPx);
      const h = this.pageCss(flow.pageHeightPx);
      const pixelRatio = this.renderPixelRatio(flow);
      this.sizeSlot(slot, w, h, index);
      if (slot.app) {
        slot.app.resize({ width: w, height: h, pixelRatio });
        this.repaint(slot.app, index);
      }
    }
  }

  private sizeSlot(slot: PageSlot, w: number, h: number, page: number): void {
    const frame = slot.el.parentElement;
    if (frame) {
      frame.style.width = `${w}px`;
      frame.style.height = `${h}px`;
      this.applyBackground(frame);
      this.applyBorders(frame, page);
      this.applyCropMarks(frame, page);
      this.applyRulers(frame, page);
    }
    slot.el.style.width = `${w}px`;
    slot.el.style.height = `${h}px`;
  }

  /** Stamp the frame's w:background — the base color (pattern fills arrive
   *  pre-averaged by the projection). */
  private applyBackground(frame: HTMLElement): void {
    const bg = this.ctx.background;
    // OOXML hex has no '#' — CSS colors do; the raw token is invalid CSS and
    // the assignment would be silently dropped.
    frame.style.backgroundColor = bg?.color ? `#${bg.color}` : "#ffffff";
    frame.style.backgroundImage = "none";
  }

  /** ST_Border tokens → CSS border-styles. Word's art borders (fancy
   *  compound/wavy lines) have no CSS counterpart and fall back to solid. */
  static readonly BORDER_STYLE: Readonly<Record<string, string>> = {
    single: "solid",
    thick: "solid",
    double: "double",
    triple: "double",
    dashed: "dashed",
    dashSmallGap: "dashed",
    dashDotStroked: "dashed",
    dotted: "dotted",
    dotDash: "dotted",
    dotDotDash: "dotted",
    threeDEmboss: "ridge",
    threeDEngrave: "groove",
    inset: "inset",
    outset: "outset",
  };

  /** The page's index within its own section (0-based) — w:pgBorders'
   *  firstPage/notFirstPage filter on the section's first page, not the
   *  document's. */
  private pageInSection(page: number): number {
    const section = this.ctx.sectionOfPage[page];
    let first = page;
    while (first > 0 && this.ctx.sectionOfPage[first - 1] === section) first--;
    return page - first;
  }

  /** Stamp a page's w:pgBorders — one absolutely-positioned div whose CSS
   *  border paints the four sides (each side carries its own style/width/
   *  color). The box insets from the page edge or the text margin per
   *  offsetFrom, scaled with the zoom like the rest of the frame. */
  private applyBorders(frame: HTMLElement, page: number): void {
    const section = this.sectionAt(page);
    const b = section.pageBorders;
    const existing = frame.querySelector(":scope > .page-borders");
    existing?.remove();
    if (!b) return;
    const nth = this.pageInSection(page);
    if (b.display === "firstPage" && nth !== 0) return;
    if (b.display === "notFirstPage" && nth === 0) return;
    const flow = section.flow;
    // offsetFrom=page measures space from the paper edge (Word's default 24
    // pt when omitted); text measures from the margin box (default 0).
    const fromPage = b.offsetFrom !== "text";
    const insetPt = (side: ProjectedPageBorder | undefined, marginPx: number): number => {
      const space = side?.spacePt ?? (fromPage ? 24 : 0);
      const px = space * (96 / 72) + (fromPage ? 0 : marginPx / this.factor);
      return px * this.factor;
    };
    const margin = {
      top: flow.contentTopPx,
      right: flow.pageWidthPx - flow.contentLeftPx - flow.contentWidthPx,
      bottom: flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx,
      left: flow.contentLeftPx,
    };
    const cssSide = (side: ProjectedPageBorder | undefined): string =>
      side
        ? `${Math.max(1, Math.round(side.widthPx * this.factor))}px ` +
          `${CanvasStage.BORDER_STYLE[side.style] ?? "solid"} ` +
          `#${side.color && side.color !== "auto" ? side.color : "000000"}`
        : "none";
    const div = document.createElement("div");
    div.className = "page-borders";
    Object.assign(div.style, {
      position: "absolute",
      pointerEvents: "none",
      // front (Word's default) paints over content; back must sit UNDER the
      // Leafer view — the view is statically positioned (z-index immune), so
      // back is negative (still above the frame's own background fill).
      zIndex: b.behind ? "-1" : "2",
    } satisfies Partial<CSSStyleDeclaration>);
    div.style.borderTop = cssSide(b.top);
    div.style.borderRight = cssSide(b.right);
    div.style.borderBottom = cssSide(b.bottom);
    div.style.borderLeft = cssSide(b.left);
    div.style.top = `${insetPt(b.top, margin.top)}px`;
    div.style.right = `${insetPt(b.right, margin.right)}px`;
    div.style.bottom = `${insetPt(b.bottom, margin.bottom)}px`;
    div.style.left = `${insetPt(b.left, margin.left)}px`;
    frame.append(div);
  }

  /** Crop marks — the four L-brackets Word draws in the margin gutter, each
   *  L's vertex just outside a content-box corner with the two 23px legs
   *  reaching into the margin. One div covers the page and carries the
   *  (zoom-scaled) margin as padding, so its content box IS the page's;
   *  eight gradient strokes offset from that origin (background-origin:
   *  content-box) land in the gutter, never over text. Leg length stays in
   *  screen px — like Word, the brackets read as fixed-size guide marks. */
  private applyCropMarks(frame: HTMLElement, page: number): void {
    const flow = this.sectionAt(page).flow;
    const pad = (px: number) => `${px * this.factor}px`;
    let div = frame.querySelector<HTMLDivElement>(":scope > .crop-marks");
    if (!div) {
      const c = "var(--docen-color-crop, #c0c0c0)";
      div = document.createElement("div");
      div.className = "crop-marks";
      Object.assign(div.style, {
        position: "absolute",
        inset: "0",
        pointerEvents: "none",
        zIndex: "2",
        backgroundOrigin: "content-box",
        backgroundRepeat: "no-repeat",
        backgroundImage: Array.from({ length: 8 }, () => `linear-gradient(${c}, ${c})`).join(", "),
        backgroundPosition: [
          "-24px -2px",
          "-2px -24px",
          "calc(100% + 24px) -2px",
          "calc(100% + 2px) -24px",
          "-24px calc(100% + 2px)",
          "-2px calc(100% + 24px)",
          "calc(100% + 24px) calc(100% + 2px)",
          "calc(100% + 2px) calc(100% + 24px)",
        ].join(", "),
        backgroundSize: [
          "23px 1px",
          "1px 23px",
          "23px 1px",
          "1px 23px",
          "23px 1px",
          "1px 23px",
          "23px 1px",
          "1px 23px",
        ].join(", "),
      } satisfies Partial<CSSStyleDeclaration>);
      frame.append(div);
    }
    div.style.padding =
      `${pad(flow.contentTopPx)} ` +
      `${pad(flow.pageWidthPx - flow.contentLeftPx - flow.contentWidthPx)} ` +
      `${pad(flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx)} ` +
      `${pad(flow.contentLeftPx)}`;
  }

  /** Rulers (Word's View → Ruler): a horizontal strip above the page and a
   *  vertical strip to its left, each an SVG of tick lines whose 0 sits on
   *  the content-box edge (Word's margin-line origin — the margin shows
   *  negative ticks). Inch ticks on en locales, centimetres otherwise; the
   *  strips re-render on every sizeSlot, so zoom rescales the ticks. They
   *  hang in the inter-page gutter (PAGE_GAP 24 > strip 20), covering
   *  nothing on the page. */
  private applyRulers(frame: HTMLElement, page: number): void {
    frame.querySelectorAll(":scope > .h-ruler, :scope > .v-ruler").forEach((el) => el.remove());
    if (!this.#showRuler) return;
    const flow = this.sectionAt(page).flow;
    const THICKNESS = 20;
    const metric = !/^en/i.test(navigator.language || "");
    const unit = (metric ? 96 / 2.54 : 96) * this.factor;
    const half = unit / 2;
    const minor = metric ? unit / 10 : unit / 4;
    const build = (length: number, zero: number, vertical: boolean): string => {
      let out = "";
      for (let p = Math.ceil(-zero / minor) * minor; p <= length - zero; p += minor) {
        const whole = p / unit;
        const major = Math.abs(whole - Math.round(whole)) < 1e-6;
        const mid = Math.abs(p / half - Math.round(p / half)) < 1e-6;
        const len = major ? THICKNESS - 2 : mid ? THICKNESS * 0.62 : THICKNESS * 0.38;
        const pos = zero + p;
        const num = Math.round(whole);
        if (vertical) {
          out += `<line x1="${THICKNESS}" y1="${pos}" x2="${THICKNESS - len}" y2="${pos}"/>`;
          if (major)
            out += `<text x="${THICKNESS - len - 2}" y="${pos + 2}" text-anchor="middle" transform="rotate(-90 ${THICKNESS - len - 2} ${pos + 2})">${num}</text>`;
        } else {
          out += `<line x1="${pos}" y1="${THICKNESS}" x2="${pos}" y2="${THICKNESS - len}"/>`;
          if (major)
            out += `<text x="${pos + 1}" y="${THICKNESS - len - 3}" stroke="none">${num}</text>`;
        }
      }
      return (
        `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%">` +
        `<g stroke="#9aa4b2" stroke-width="1" fill="#5b6675" font-size="7"` +
        ` font-family="Inter, sans-serif">${out}</g></svg>`
      );
    };
    const mount = (cls: string, style: Partial<CSSStyleDeclaration>, svg: string): void => {
      const div = document.createElement("div");
      div.className = cls;
      // Two assign targets, not one spread object: the linter flags spreading
      // a CSSStyleDeclaration-typed value (index-signature interface) into an
      // object literal as an iterable spread.
      Object.assign(
        div.style,
        {
          position: "absolute",
          pointerEvents: "none",
          zIndex: "2",
          background: "#fafbfc",
          border: "1px solid #d8dce2",
        } satisfies Partial<CSSStyleDeclaration>,
        style,
      );
      div.innerHTML = svg;
      frame.append(div);
    };
    mount(
      "h-ruler",
      {
        left: "0",
        top: `-${THICKNESS}px`,
        width: `${this.pageCss(flow.pageWidthPx)}px`,
        height: `${THICKNESS}px`,
      },
      build(this.pageCss(flow.pageWidthPx), flow.contentLeftPx * this.factor, false),
    );
    mount(
      "v-ruler",
      {
        left: `-${THICKNESS}px`,
        top: "0",
        width: `${THICKNESS}px`,
        height: `${this.pageCss(flow.pageHeightPx)}px`,
      },
      build(this.pageCss(flow.pageHeightPx), flow.contentTopPx * this.factor, true),
    );
  }

  /** Lay out page slots for a flow result and repaint visible pages. The
   *  stage is built once and lives across documents — every sync must
   *  refresh the context (an opened file's headers/footers arrive here).
   *  Multi-section documents pass one {@link CanvasStageSection} per section
   *  plus the page→section map; a single-section document is a one-entry
   *  list. `dirty` (parallel to `pages`, absent = all) marks the pages whose
   *  layout changed — the others keep their painted canvas untouched, which
   *  is what makes a one-word edit cost one repaint instead of ninety-one. */
  sync(
    pages: FlowPage[],
    sections: CanvasStageSection[],
    sectionOfPage: number[],
    background?: ProjectedPageBackground,
    dirty?: readonly boolean[],
  ): void {
    this.pages = pages;
    this.ctx.sections = sections;
    this.ctx.sectionOfPage = sectionOfPage;
    // A pure derived value (recomputed per render) — always overwritten so a
    // document without a background clears the previous one's tile.
    this.ctx.background = background;
    // Line numbers count across pages (continuous runs through page breaks)
    // — one pass over the whole flow, keyed by page for the per-page paint.
    this.lineNumberMarks = computeLineNumbers(pages, sections, sectionOfPage);

    while (this.slots.length < pages.length) {
      // New slots clone the size of the section their page belongs to.
      const flow = this.sectionAt(this.slots.length).flow;
      const w = this.pageCss(flow.pageWidthPx);
      const h = this.pageCss(flow.pageHeightPx);
      const frame = document.createElement("div");
      Object.assign(frame.style, {
        position: "relative",
        width: `${w}px`,
        height: `${h}px`,
        marginBottom: `${PAGE_GAP}px`,
        boxShadow: "0 1px 3px rgba(0,0,0,.2), 0 4px 12px rgba(0,0,0,.08)",
      } satisfies Partial<CSSStyleDeclaration>);
      // Pristine inner div the App takes over as its view.
      const el = document.createElement("div");
      el.style.width = `${w}px`;
      el.style.height = `${h}px`;
      frame.append(el);
      this.shell.append(frame);
      this.slots.push({ el, app: null, layers: null });
      this.io.observe(el);
    }
    while (this.slots.length > pages.length) {
      const slot = this.slots.pop()!;
      this.io.unobserve(slot.el);
      slot.app?.destroy();
      slot.el.parentElement?.remove();
    }
    // A zoom applied between syncs (initial attr → first sync) or a section
    // mix re-sizes created slots to their page's own section.
    for (const [index, slot] of this.slots.entries()) {
      const flow = this.sectionAt(index).flow;
      this.sizeSlot(slot, this.pageCss(flow.pageWidthPx), this.pageCss(flow.pageHeightPx), index);
    }
    for (const [index, slot] of this.slots.entries()) {
      // An absent `dirty` is the caller's structural signal (section
      // geometry/furniture/background changed) — repaint flat. With a dirty
      // array every repainted page is a body-only change, so its furniture
      // layers survive (see repaint).
      if (slot.app && dirty?.[index] !== false) {
        this.repaint(slot.app, index, dirty != null);
      }
    }
  }

  /** Which furniture slot a page displays — the edit story's data source
   *  (an absent first/even slot falls back to default at pick time). */
  slotOfPage(page: number): "default" | "first" | "even" {
    const slot = this.slotOf(page);
    return slot === 1 ? "first" : slot === 2 ? "even" : "default";
  }

  /** A section's slot stack for a page (an absent slot falls back to
   *  default). Null when the doc has none. */
  private slotStackOf(kind: "header" | "footer", page: number): LaidFurnitureSlot | undefined {
    const laid = this.sectionAt(page).furnitureLaid;
    const slots = laid?.[kind];
    return slots?.[this.slotOf(page)] ?? slots?.[0];
  }

  /** The laid furniture stack a page displays (its section's stacks; an
   *  absent slot falls back to default). Null when the doc has none. */
  furnitureStack(kind: "header" | "footer", page = 0): readonly LaidOutStackItem[] | null {
    return this.slotStackOf(kind, page)?.stack ?? null;
  }

  /** A page's editable furniture band, page-local — [top, bottom) with the
   *  dashed boundary at the inner edge (bottom for headers, top for
   *  footers). Two heights: `paintY` anchors the caret map at the stack's
   *  actual draw y (the painter uses the raw stack height — no floor), while
   *  the band extent is floored at one strut line so an empty story is
   *  still enterable (hit target + dashed boundary). */
  furnitureBand(
    kind: "header" | "footer",
    page = 0,
  ): { top: number; bottom: number; paintY: number } | null {
    const { flow, furniture: f } = this.sectionAt(page);
    if (!f) return null;
    const stackH = this.slotStackOf(kind, page)?.heightPx ?? 0;
    const bandH = Math.max(stackH, 24);
    if (kind === "header") {
      const paintY = f.headerDistancePx ?? 48;
      return { top: 0, bottom: paintY + bandH, paintY };
    }
    const paintY = flow.pageHeightPx - (f.footerDistancePx ?? 48) - stackH;
    return { top: paintY - (bandH - stackH), bottom: flow.pageHeightPx, paintY };
  }

  /** Story chrome while a header/footer is being edited: Word grays the body
   *  and draws the dashed boundary + a gray "Header"/"Footer" tag on every
   *  page (the boundary is an interaction affordance, never printed). */
  setStoryEdit(edit: { kind: "header" | "footer"; label: string } | null): void {
    this.storyEdit = edit;
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
    }
  }

  private storyEdit: { kind: "header" | "footer"; label: string } | null = null;

  /** A page's slot element (scrollIntoView target for page jumps). */
  slotAt(index: number): HTMLElement | null {
    return this.slots[index]?.el ?? null;
  }

  /** Rasterize every page for printing: pages the IntersectionObserver never
   *  reached get their App forced (a printout needs all pages, not just the
   *  scrolled-into-view ones), every slot repaints and force-renders, then
   *  each canvas exports as PNG. `width`/`height` are the page's unzoomed CSS
   *  px (96 dpi) so the print view can lay the images out at true paper size. */
  printSnapshots(): { width: number; height: number; url: string }[] {
    for (const [index, slot] of this.slots.entries()) {
      this.ensure(slot);
      if (!slot.app) continue;
      this.repaint(slot.app, index);
      slot.app.forceRender();
    }
    const shots: { width: number; height: number; url: string }[] = [];
    for (const [index, slot] of this.slots.entries()) {
      const canvas = slot.el.querySelector("canvas");
      if (!canvas) continue;
      const flow = this.sectionAt(index).flow;
      shots.push({
        width: this.pageCss(flow.pageWidthPx),
        height: this.pageCss(flow.pageHeightPx),
        url: canvas.toDataURL("image/png"),
      });
    }
    return shots;
  }

  destroy(): void {
    this.io.disconnect();
    this.dprMedia?.removeEventListener("change", this.dprChange);
    for (const slot of this.slots) slot.app?.destroy();
    releasePinnedImages();
    this.shell.remove();
  }

  private ensure(slot: PageSlot): void {
    if (slot.app) return;
    slot.layers = null;
    const app = new App({
      view: slot.el,
      fill: "transparent",
      // Explicit DPR (Leafer's default samples it at creation anyway) so the
      // value stays consistent across app.resize calls.
      pixelRatio: this.renderPixelRatio(this.sectionAt(this.slots.indexOf(slot)).flow),
      editor: { moveable: false },
      // The document surface is MS Office-shaped: scrolling belongs to the
      // outer DOM container, never to the canvas. `leafer-editor` (imported by
      // the picture surface) registers viewport plugins globally which would
      // otherwise pan the App's world on wheel — disable both groups.
      move: { disabled: true },
      wheel: { disabled: true },
    });
    if (app.editor) app.editor.visible = false;
    slot.app = app;
    this.repaint(app, this.slots.indexOf(slot));
  }

  /** Which slot a page uses: first (titlePage) on its SECTION's first page,
   *  even on even PHYSICAL pages (the pattern runs document-wide — Word's
   *  evenAndOddHeaders is a settings-level switch), else default. */
  private slotOf(index: number): number {
    const section = this.sectionAt(index);
    const f = section.furniture;
    const local = index - this.ctx.sectionOfPage.indexOf(this.ctx.sectionOfPage[index] ?? 0);
    if (local === 0 && f?.titlePage) return 1;
    if ((index + 1) % 2 === 0 && f?.evenAndOddHeaders) return 2;
    return 0;
  }

  /** Repaint one page. `keepFurniture` (the per-keystroke body-only path)
   *  rebuilds just the body-bearing layers in place, preserving the furniture
   *  groups — headers/footers are per-section constants in Word, and
   *  re-creating their image leaves a decode gap on every keystroke that
   *  reads as header flicker while deleting. Any structural change (section
   *  geometry, furniture data, background) or a story edit repaints flat. */
  private repaint(app: App, index: number, keepFurniture = false): void {
    // app.tree is an ILeafer (extends IGroup) — clear/add come with it.
    const tree = app.tree;
    if (!tree) return;
    // The canvas is laid out at the zoomed size; all paint coordinates below
    // stay in unzoomed page px and this scale maps them onto the bitmap.
    tree.scale = this.factor;
    // This page paints with its OWN section's box + furniture.
    const { flow, furniture, lineNumbers, columns } = this.sectionAt(index);
    const marks = this.lineNumberMarks.get(index);
    const ctx: PaintContext = {
      metrics: this.ctx.metrics,
      flow,
      furniture,
      background: this.ctx.background,
      pageIndex: index,
      pageCount: this.pages.length,
      layer: "behind",
      showMarks: this.#showMarks,
      showGridlines: this.#showGridlines,
      marksLabels: this.ctx.marksLabels,
      // Async image decodes landing after this repaint need the same eager
      // render — the change-driven scheduler cannot be relied on here.
      rerender: () => app.forceRender(),
      // In-front floats collect here through the body pass and paint after
      // its last paragraph (Word stacks them above ALL text).
      deferredDrawings: [],
      columns,
      ...(lineNumbers && marks?.length ? { lineNumbers: { config: lineNumbers, marks } } : {}),
    };
    const slotIndex = this.slotOf(index);
    const items = this.pages[index]?.items ?? [];
    // The body pass also catalogs every drawing's box for click hit-testing
    // (behind-doc floats included — the earlier pass painted them, this one
    // records where).
    const hitBoxes: DrawingHitBox[] = [];

    const layers = keepFurniture && !this.storyEdit ? this.slots[index]!.layers : null;
    if (layers) {
      ctx.hitBoxes = hitBoxes;
      layers.behind.clear();
      paintScene(layers.behind, items, ctx);
      paintGridlines(layers.behind, ctx);
      ctx.layer = "body";
      layers.body.clear();
      paintScene(layers.body, items, ctx);
      paintLineNumbers(layers.body, ctx);
      paintColumnSeparators(layers.body, ctx);
      paintFootnotes(layers.body, this.pages[index]?.footnotes, ctx);
      this.#flushDrawings(ctx);
      app.forceRender();
      this.hitBoxes.set(index, hitBoxes);
      return;
    }

    tree.clear();
    this.slots[index]!.layers = null;
    // The page's w:background must tint the BITMAP itself (exports and pixel
    // probes read the canvas, not the frame's CSS), so the base color paints
    // as the bottommost scene element. The App-level `fill` cannot host it:
    // an App has no canvas of its own and drops `fill` from the child-layer
    // configs it builds (App.ts __getChildConfig).
    const bg = this.ctx.background;
    if (bg?.color) {
      tree.add(
        new Rect({
          x: 0,
          y: 0,
          width: flow.pageWidthPx,
          height: flow.pageHeightPx,
          fill: `#${bg.color}`,
          hittable: false,
        }),
      );
    }
    // Word's stacking: behind-text floats sit under everything from the text
    // layer — footer furniture included (a full-bleed backdrop never covers
    // the page number). Paint the behind pass first, then furniture + the
    // body pass on top; in-front floats close the sequence inside the body
    // pass's paragraphs. While a furniture story is being edited the body
    // pass sits under a white veil instead and the furniture paints on top
    // of it (the story under edit stays fully opaque) — the story-edit order
    // interleaves body and furniture, so it paints flat (no layers).
    if (!this.storyEdit) {
      const pageLayers: PageLayers = {
        behind: new Group(),
        furnitureBehind: new Group(),
        furnitureBody: new Group(),
        body: new Group(),
      };
      tree.add([
        pageLayers.behind,
        pageLayers.furnitureBehind,
        pageLayers.furnitureBody,
        pageLayers.body,
      ]);
      paintScene(pageLayers.behind, items, ctx);
      paintGridlines(pageLayers.behind, ctx);
      ctx.layer = "body";
      ctx.hitBoxes = hitBoxes;
      this.paintFurniture(
        pageLayers.furnitureBehind,
        pageLayers.furnitureBody,
        slotIndex,
        ctx,
        flow,
        furniture,
      );
      paintScene(pageLayers.body, items, ctx);
      paintLineNumbers(pageLayers.body, ctx);
      paintColumnSeparators(pageLayers.body, ctx);
      paintFootnotes(pageLayers.body, this.pages[index]?.footnotes, ctx);
      this.#flushDrawings(ctx);
      this.slots[index]!.layers = pageLayers;
    } else {
      ctx.hitBoxes = hitBoxes;
      paintScene(tree, items, ctx);
      paintLineNumbers(tree, ctx);
      paintColumnSeparators(tree, ctx);
      paintFootnotes(tree, this.pages[index]?.footnotes, ctx);
      // Under the story-edit veil like the rest of the body — the story being
      // edited paints above (opaque) on top.
      this.#flushDrawings(ctx);
      tree.add(
        new Rect({
          x: 0,
          y: 0,
          width: flow.pageWidthPx,
          height: flow.pageHeightPx,
          fill: "rgba(255,255,255,.62)",
          hittable: false,
        }),
      );
      this.paintFurniture(tree, tree, slotIndex, ctx, flow, furniture);
      this.paintStoryBoundary(tree, ctx, flow);
    }
    // Render eagerly: Leafer's change-driven scheduling stalls when the App
    // was created while its view was offscreen (an IO callback during mount)
    // and never picks the page back up.
    app.forceRender();
    this.hitBoxes.set(index, hitBoxes);
  }

  /** Paint the floats both passes parked in the queue — after the pass's
   *  last paragraph, so no text paragraph can paint over them (Word stacks
   *  in-front floats above ALL body text). Each band sorts by
   *  w:relativeHeight first: same-band stacking follows the z-order, ties
   *  keep document order (a stable sort). */
  #flushDrawings(ctx: PaintContext): void {
    const queue = ctx.deferredDrawings;
    if (!queue) return;
    for (const layer of ["behind", "body"] as const) {
      const band = queue.filter((entry) => entry.layer === layer).sort((a, b) => a.z - b.z);
      if (band.length === 0) continue;
      ctx.layer = layer;
      for (const entry of band) entry.paint();
    }
    queue.length = 0;
  }

  /** The page's drawing boxes as the body pass painted them — the click
   *  hit table (empty until the page repaints at least once). */
  private readonly hitBoxes = new Map<number, DrawingHitBox[]>();

  /** The topmost drawing whose painted box contains the page-local point
   *  (null when none does) — later-painted wins, Word's z-click. */
  drawingAt(page: number, lx: number, ly: number): DrawingHitBox | null {
    const boxes = this.hitBoxes.get(page);
    if (!boxes) return null;
    for (let i = boxes.length - 1; i >= 0; i--) {
      const b = boxes[i]!;
      if (lx >= b.x && lx <= b.x + b.width && ly >= b.y && ly <= b.y + b.height) return b;
    }
    return null;
  }

  /** The painted box of a paragraph's index-th drawing across pages — the
   *  selection overlay's geometry source (a re-render may have moved the
   *  host paragraph, and with it the drawing, onto another page). */
  drawingBoxOf(para: unknown, index: number, kind: DrawingHitBox["kind"]): DrawingHitBox | null {
    for (const boxes of this.hitBoxes.values()) {
      const b = boxes.find((box) => box.para === para && box.index === index && box.kind === kind);
      if (b) return b;
    }
    return null;
  }

  /** Both furniture stacks for a page's slot, at their page positions —
   *  distances come from the page's own section. The paint context clears
   *  the body flow's grid pitch: the header/footer story keeps natural line
   *  heights (its paragraphs were laid out with no grid context either), so
   *  a text box inside the story must not inherit the body's docGrid. */
  private paintFurniture(
    behind: IGroup,
    body: IGroup,
    slot: number,
    ctx: PaintContext,
    flow: ProjectedFlowBox,
    furniture: ProjectedPageFurniture | undefined,
  ): void {
    const storyFlow = flow.linePitchPx ? { ...flow, linePitchPx: undefined } : flow;
    const section = this.sectionAt(ctx.pageIndex);
    const header = section.furnitureLaid?.header[slot] ?? section.furnitureLaid?.header[0];
    const footer = section.furnitureLaid?.footer[slot] ?? section.furnitureLaid?.footer[0];
    const paintSlots = (layer: PaintContext["layer"]): void => {
      const storyCtx: PaintContext = { ...ctx, flow: storyFlow, layer };
      if (header) {
        paintFurnitureStack(
          layer === "behind" ? behind : body,
          header.stack,
          flow.contentLeftPx,
          furniture?.headerDistancePx ?? 48,
          storyCtx,
        );
      }
      if (footer) {
        const bottom = flow.pageHeightPx - (furniture?.footerDistancePx ?? 48);
        paintFurnitureStack(
          layer === "behind" ? behind : body,
          footer.stack,
          flow.contentLeftPx,
          bottom - footer.heightPx,
          storyCtx,
        );
      }
    };
    // Furniture paragraphs carry anchored drawings through the same
    // projection as body paragraphs — Word's header-anchored watermarks are
    // behind-doc shapes. Paint those first, beneath the body's own behind
    // floats and just above the page background; the body pass paints the
    // story text and front drawings as before.
    paintSlots("behind");
    paintSlots("body");
  }

  /** The dashed boundary of the story under edit plus its gray tag — a
   *  header's boundary runs under its content, a footer's runs above it;
   *  the tag sits just above the line in both cases (Word's layout). */
  private paintStoryBoundary(tree: IGroup, ctx: PaintContext, flow: ProjectedFlowBox): void {
    const edit = this.storyEdit;
    if (!edit) return;
    const band = this.furnitureBand(edit.kind, ctx.pageIndex);
    if (!band) return;
    const dashY = edit.kind === "header" ? band.bottom : band.top;
    tree.add(
      new Line({
        points: [flow.contentLeftPx, dashY, flow.contentLeftPx + flow.contentWidthPx, dashY],
        stroke: "#a6a6a6",
        strokeWidth: 1,
        dashPattern: [4, 3, 1, 3],
        hittable: false,
      }),
    );
    tree.add(
      new Text({
        x: flow.contentLeftPx,
        y: dashY - 17,
        text: edit.label,
        fill: "#8a8886",
        fontSize: 11,
        hittable: false,
      }),
    );
  }
}
