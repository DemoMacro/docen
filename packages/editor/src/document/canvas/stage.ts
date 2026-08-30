import {
  paintFurnitureStack,
  paintScene,
  releasePinnedImages,
  type DrawingHitBox,
  type PaintContext,
} from "@docen/core/painter";
import type {
  ProjectedFlowBox,
  ProjectedPageBackground,
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
import { App, Line, Rect, Text, type IGroup } from "leafer-ui";

const PAGE_GAP = 24;

/** One laid furniture slot: the paint-ready stack and its laid height. */
export interface LaidFurnitureSlot {
  stack: readonly LaidOutStackItem[];
  heightPx: number;
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
  /** Headers/footers for this section's pages (absent = none). */
  furniture?: ProjectedPageFurniture;
  /** The slots of `furniture` laid out once (layFurnitureSections) — the
   *  page insets push the body by these heights and the painter draws these
   *  same stacks, so push-down == painted band height by construction. */
  furnitureLaid?: LaidFurnitureSection;
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
}

export class CanvasStage {
  readonly shell: HTMLElement;
  private readonly slots: { el: HTMLElement; app: App | null }[] = [];
  private readonly io: IntersectionObserver;
  private pages: FlowPage[] = [];

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
          }
        }
      },
      // The IO root is the SCROLL CONTAINER (the stage host's parent), not the
      // viewport: rootMargin only widens the root's clip rect, and an
      // intermediate overflow:auto box re-clips it — pages below the scroller's
      // visible area would never pre-render against a viewport root. The bottom
      // margin (order top/right/bottom/left) pre-renders ahead of the scroll
      // direction so fast scrolls rarely meet a blank page.
      { root: stage.parentElement, rootMargin: "0px 0px 150% 0px" },
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

  private applyZoom(): void {
    // Sections can carry different paper sizes — each slot sizes to its own.
    for (const [index, slot] of this.slots.entries()) {
      const flow = this.sectionAt(index).flow;
      const w = this.pageCss(flow.pageWidthPx);
      const h = this.pageCss(flow.pageHeightPx);
      const pixelRatio = this.renderPixelRatio(flow);
      this.sizeSlot(slot, w, h);
      if (slot.app) {
        slot.app.resize({ width: w, height: h, pixelRatio });
        this.repaint(slot.app, index);
      }
    }
  }

  private sizeSlot(slot: { el: HTMLElement; app: App | null }, w: number, h: number): void {
    const frame = slot.el.parentElement;
    if (frame) {
      frame.style.width = `${w}px`;
      frame.style.height = `${h}px`;
      this.applyBackground(frame);
    }
    slot.el.style.width = `${w}px`;
    slot.el.style.height = `${h}px`;
  }

  /** Stamp the frame's w:background — base color plus the pattern tile sized
   *  to the current zoom (CSS background scales with the page, so the pattern
   *  reads identically at every level). */
  private applyBackground(frame: HTMLElement): void {
    const bg = this.ctx.background;
    // OOXML hex has no '#' — CSS colors do; the raw token is invalid CSS and
    // the assignment would be silently dropped.
    frame.style.backgroundColor = bg?.color ? `#${bg.color}` : "#ffffff";
    if (bg?.tileSrc && bg.tilePx) {
      const size = Math.max(1, Math.round(bg.tilePx * this.factor));
      frame.style.backgroundImage = `url(${bg.tileSrc})`;
      frame.style.backgroundSize = `${size}px ${size}px`;
    } else {
      frame.style.backgroundImage = "none";
    }
  }

  /** Lay out page slots for a flow result and repaint visible pages. The
   *  stage is built once and lives across documents — every sync must
   *  refresh the context (an opened file's headers/footers arrive here).
   *  Multi-section documents pass one {@link CanvasStageSection} per section
   *  plus the page→section map; a single-section document is a one-entry
   *  list. */
  sync(
    pages: FlowPage[],
    sections: CanvasStageSection[],
    sectionOfPage: number[],
    background?: ProjectedPageBackground,
  ): void {
    this.pages = pages;
    this.ctx.sections = sections;
    this.ctx.sectionOfPage = sectionOfPage;
    // A pure derived value (recomputed per render) — always overwritten so a
    // document without a background clears the previous one's tile.
    this.ctx.background = background;

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
      this.slots.push({ el, app: null });
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
      this.sizeSlot(slot, this.pageCss(flow.pageWidthPx), this.pageCss(flow.pageHeightPx));
    }
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
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
   *  footers); `paintY` is the stack's own draw y (the band reaches past it
   *  to the page edge, the caret map must anchor at the stack itself). At
   *  least one strut line tall so an empty story is enterable. */
  furnitureBand(
    kind: "header" | "footer",
    page = 0,
  ): { top: number; bottom: number; paintY: number } | null {
    const { flow, furniture: f } = this.sectionAt(page);
    if (!f) return null;
    const h = Math.max(this.slotStackOf(kind, page)?.heightPx ?? 0, 24);
    if (kind === "header") {
      const paintY = f.headerDistancePx ?? 48;
      return { top: 0, bottom: paintY + h, paintY };
    }
    const paintY = flow.pageHeightPx - (f.footerDistancePx ?? 48) - h;
    return { top: paintY, bottom: flow.pageHeightPx, paintY };
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

  destroy(): void {
    this.io.disconnect();
    this.dprMedia?.removeEventListener("change", this.dprChange);
    for (const slot of this.slots) slot.app?.destroy();
    releasePinnedImages();
    this.shell.remove();
  }

  private ensure(slot: { el: HTMLElement; app: App | null }): void {
    if (slot.app) return;
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

  private repaint(app: App, index: number): void {
    // app.tree is an ILeafer (extends IGroup) — clear/add come with it.
    const tree = app.tree;
    if (!tree) return;
    // The canvas is laid out at the zoomed size; all paint coordinates below
    // stay in unzoomed page px and this scale maps them onto the bitmap.
    tree.scale = this.factor;
    tree.clear();
    // This page paints with its OWN section's box + furniture.
    const { flow, furniture } = this.sectionAt(index);
    const ctx: PaintContext = {
      metrics: this.ctx.metrics,
      flow,
      furniture,
      background: this.ctx.background,
      pageIndex: index,
      pageCount: this.pages.length,
      layer: "behind",
      // Async image decodes landing after this repaint need the same eager
      // render — the change-driven scheduler cannot be relied on here.
      rerender: () => app.forceRender(),
    };
    const slot = this.slotOf(index);
    // Word's stacking: behind-text floats sit under everything from the text
    // layer — footer furniture included (a full-bleed backdrop never covers
    // the page number). Paint the behind pass first, then furniture + the
    // body pass on top; in-front floats close the sequence inside the body
    // pass's paragraphs. While a furniture story is being edited the body
    // pass sits under a white veil instead and the furniture paints on top
    // of it (the story under edit stays fully opaque).
    const items = this.pages[index]?.items ?? [];
    ctx.layer = "behind";
    // The page's w:background must tint the BITMAP itself (exports and pixel
    // probes read the canvas, not the frame's CSS), so the base color / tile
    // pattern paints as the bottommost scene element. The App-level `fill`
    // cannot host it: an App has no canvas of its own and drops `fill` from
    // the child-layer configs it builds (App.ts __getChildConfig).
    const bg = this.ctx.background;
    if (bg) {
      tree.add(
        new Rect({
          x: 0,
          y: 0,
          width: flow.pageWidthPx,
          height: flow.pageHeightPx,
          fill:
            bg.tileSrc && bg.tilePx
              ? { type: "image", url: bg.tileSrc, repeat: true, size: bg.tilePx }
              : bg.color
                ? `#${bg.color}`
                : undefined,
          hittable: false,
        }),
      );
    }
    paintScene(tree, items, ctx);
    ctx.layer = "body";
    // The body pass also catalogs every drawing's box for click hit-testing
    // (behind-doc floats included — the earlier pass painted them, this one
    // records where).
    const hitBoxes: DrawingHitBox[] = [];
    ctx.hitBoxes = hitBoxes;
    if (!this.storyEdit) {
      this.paintFurniture(tree, slot, ctx, flow, furniture);
      paintScene(tree, items, ctx);
    } else {
      paintScene(tree, items, ctx);
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
      this.paintFurniture(tree, slot, ctx, flow, furniture);
      this.paintStoryBoundary(tree, ctx, flow);
    }
    // Render eagerly: Leafer's change-driven scheduling stalls when the App
    // was created while its view was offscreen (an IO callback during mount)
    // and never picks the page back up.
    app.forceRender();
    this.hitBoxes.set(index, hitBoxes);
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
    tree: IGroup,
    slot: number,
    ctx: PaintContext,
    flow: ProjectedFlowBox,
    furniture: ProjectedPageFurniture | undefined,
  ): void {
    ctx.layer = "body";
    const storyFlow = flow.linePitchPx ? { ...flow, linePitchPx: undefined } : flow;
    const storyCtx: PaintContext = { ...ctx, flow: storyFlow };
    const section = this.sectionAt(ctx.pageIndex);
    const header = section.furnitureLaid?.header[slot] ?? section.furnitureLaid?.header[0];
    if (header) {
      paintFurnitureStack(
        tree,
        header.stack,
        flow.contentLeftPx,
        furniture?.headerDistancePx ?? 48,
        storyCtx,
      );
    }
    const footer = section.furnitureLaid?.footer[slot] ?? section.furnitureLaid?.footer[0];
    if (footer) {
      const bottom = flow.pageHeightPx - (furniture?.footerDistancePx ?? 48);
      paintFurnitureStack(
        tree,
        footer.stack,
        flow.contentLeftPx,
        bottom - footer.heightPx,
        storyCtx,
      );
    }
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
