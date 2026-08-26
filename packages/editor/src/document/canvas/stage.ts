import type { ProjectedFlowBox, ProjectedPageFurniture } from "@docen/docx/layout";
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
import { App } from "leafer-ui";

import { paintFurnitureStack, paintScene, type PaintContext } from "./scene";

const PAGE_GAP = 24;

/** Font metrics (baseline ratios for half-leading) + page geometry. */
export interface CanvasStageContext {
  metrics: FontMetrics;
  flow: ProjectedFlowBox;
  /** Headers/footers to paint on every page (absent = none). */
  furniture?: ProjectedPageFurniture;
}

export class CanvasStage {
  readonly shell: HTMLElement;
  private readonly slots: { el: HTMLElement; app: App | null }[] = [];
  private readonly io: IntersectionObserver;
  private pages: FlowPage[] = [];
  /** Header/footer stacks per furniture slot, laid out once per sync at the
   *  content width (furniture never reflows across pages). */
  private furnitureStacks: Map<readonly LaidOutStackItem[], number> = new Map();

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

  private renderPixelRatio(): number {
    const w = this.ctx.flow.pageWidthPx * this.factor;
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
    if (this.slots.length === 0) return;
    const w = this.pageCss(this.ctx.flow.pageWidthPx);
    const h = this.pageCss(this.ctx.flow.pageHeightPx);
    const pixelRatio = this.renderPixelRatio();
    for (const [index, slot] of this.slots.entries()) {
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
    }
    slot.el.style.width = `${w}px`;
    slot.el.style.height = `${h}px`;
  }

  /** Lay out page slots for a flow result and repaint visible pages. The
   *  stage is built once and lives across documents — every sync must
   *  refresh the context (an opened file's headers/footers arrive here). */
  sync(pages: FlowPage[], flow: ProjectedFlowBox, furniture?: ProjectedPageFurniture): void {
    this.pages = pages;
    this.ctx.flow = flow;
    if (furniture !== undefined) this.ctx.furniture = furniture;
    this.layoutFurniture();
    const w = this.pageCss(flow.pageWidthPx);
    const h = this.pageCss(flow.pageHeightPx);

    while (this.slots.length < pages.length) {
      const frame = document.createElement("div");
      Object.assign(frame.style, {
        position: "relative",
        width: `${w}px`,
        height: `${h}px`,
        marginBottom: `${PAGE_GAP}px`,
        backgroundColor: "#ffffff",
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
    // A zoom applied between syncs (initial attr → first sync) re-sizes
    // already-created slots too.
    for (const slot of this.slots) this.sizeSlot(slot, w, h);
    while (this.slots.length > pages.length) {
      const slot = this.slots.pop()!;
      this.io.unobserve(slot.el);
      slot.app?.destroy();
      slot.el.parentElement?.remove();
    }
    for (const [index, slot] of this.slots.entries()) {
      if (slot.app) this.repaint(slot.app, index);
    }
  }

  /** A page's slot element (scrollIntoView target for page jumps). */
  slotAt(index: number): HTMLElement | null {
    return this.slots[index]?.el ?? null;
  }

  destroy(): void {
    this.io.disconnect();
    this.dprMedia?.removeEventListener("change", this.dprChange);
    for (const slot of this.slots) slot.app?.destroy();
    this.shell.remove();
  }

  private ensure(slot: { el: HTMLElement; app: App | null }): void {
    if (slot.app) return;
    const app = new App({
      view: slot.el,
      fill: "transparent",
      // Explicit DPR (Leafer's default samples it at creation anyway) so the
      // value stays consistent across app.resize calls.
      pixelRatio: this.renderPixelRatio(),
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

  /** Lay each furniture slot out once (width = the content box, the same grid
   *  context as body blocks) and remember each stack's height for the footer's
   *  bottom-edge placement. */
  private layoutFurniture(): void {
    this.furnitureStacks = new Map();
    const f = this.ctx.furniture;
    if (!f) return;
    const measurer = new TextMeasurer(this.ctx.metrics);
    const ctx = this.ctx.flow.linePitchPx ? { linePitchPx: this.ctx.flow.linePitchPx } : undefined;
    const lay = (blocks: typeof f.header): readonly LaidOutStackItem[] | undefined => {
      if (!blocks) return undefined;
      const laid = stackBlocks(blocks, this.ctx.flow.contentWidthPx, ctx, measurer);
      this.furnitureStacks.set(laid.stack, laid.heightPx);
      return laid.stack;
    };
    this.headerStacks = [lay(f.header), lay(f.firstHeader), lay(f.evenHeader)];
    this.footerStacks = [lay(f.footer), lay(f.firstFooter), lay(f.evenFooter)];
  }

  /** Per-slot header/footer stacks [default, first, even] — an undefined slot
   *  falls back to default at pick time (OOXML). */
  private headerStacks: (readonly LaidOutStackItem[] | undefined)[] = [];
  private footerStacks: (readonly LaidOutStackItem[] | undefined)[] = [];

  /** Which slot a page uses: first (titlePage) on page 1, even on even pages
   *  when the document asks for different even/odd headers, else default. */
  private slotOf(index: number): number {
    const f = this.ctx.furniture;
    if (index === 0 && f?.titlePage) return 1;
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
    const ctx: PaintContext = {
      ...this.ctx,
      pageIndex: index,
      pageCount: this.pages.length,
    };
    const { flow } = this.ctx;
    const slot = this.slotOf(index);
    const header = this.headerStacks[slot] ?? this.headerStacks[0];
    if (header) {
      paintFurnitureStack(
        tree,
        header,
        flow.contentLeftPx,
        this.ctx.furniture?.headerDistancePx ?? 48,
        ctx,
      );
    }
    const footer = this.footerStacks[slot] ?? this.footerStacks[0];
    if (footer) {
      const height = this.furnitureStacks.get(footer) ?? 0;
      const bottom = flow.pageHeightPx - (this.ctx.furniture?.footerDistancePx ?? 48);
      paintFurnitureStack(tree, footer, flow.contentLeftPx, bottom - height, ctx);
    }
    paintScene(tree, this.pages[index]?.items ?? [], ctx);
    // Render eagerly: Leafer's change-driven scheduling stalls when the App
    // was created while its view was offscreen (an IO callback during mount)
    // and never picks the page back up.
    app.forceRender();
  }
}
