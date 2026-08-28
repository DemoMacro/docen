import type {
  ProjectedFlowBox,
  ProjectedPageBackground,
  ProjectedPageFurniture,
} from "@docen/docx/layout";
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

import { paintFurnitureStack, paintScene, type PaintContext } from "./scene";

const PAGE_GAP = 24;

/** Font metrics (baseline ratios for half-leading) + page geometry. */
export interface CanvasStageContext {
  metrics: FontMetrics;
  flow: ProjectedFlowBox;
  /** Headers/footers to paint on every page (absent = none). */
  furniture?: ProjectedPageFurniture;
  /** Page background (w:background — base color + optional pattern tile). */
  background?: ProjectedPageBackground;
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
    frame.style.backgroundColor = bg?.color ?? "#ffffff";
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
   *  refresh the context (an opened file's headers/footers arrive here). */
  sync(
    pages: FlowPage[],
    flow: ProjectedFlowBox,
    furniture?: ProjectedPageFurniture,
    background?: ProjectedPageBackground,
  ): void {
    this.pages = pages;
    this.ctx.flow = flow;
    if (furniture !== undefined) this.ctx.furniture = furniture;
    // A pure derived value (recomputed per render) — always overwritten so a
    // document without a background clears the previous one's tile.
    this.ctx.background = background;
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

  /** Re-lay the furniture stacks and repaint every live page — the
   *  header/footer story's render path. The body flow is untouched: pages and
   *  their laid blocks stay, only the furniture trees rebuild. */
  syncFurniture(furniture?: ProjectedPageFurniture, background?: ProjectedPageBackground): void {
    if (furniture !== undefined) this.ctx.furniture = furniture;
    if (background !== undefined) this.ctx.background = background;
    this.layoutFurniture();
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

  /** The laid furniture stack a page displays (its geometry, whether the
   *  stack's own slot or the default fallback). Null when the doc has none. */
  furnitureStack(kind: "header" | "footer", page = 0): readonly LaidOutStackItem[] | null {
    if (!this.ctx.furniture) return null;
    const slot = this.slotOf(page);
    const stacks = kind === "header" ? this.headerStacks : this.footerStacks;
    return stacks[slot] ?? stacks[0] ?? null;
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
    if (!this.ctx.furniture) return null;
    const f = this.ctx.furniture;
    const stack = this.furnitureStack(kind, page);
    const h = Math.max(stack ? (this.furnitureStacks.get(stack) ?? 0) : 0, 24);
    if (kind === "header") {
      const paintY = f.headerDistancePx ?? 48;
      return { top: 0, bottom: paintY + h, paintY };
    }
    const paintY = this.ctx.flow.pageHeightPx - (f.footerDistancePx ?? 48) - h;
    return { top: paintY, bottom: this.ctx.flow.pageHeightPx, paintY };
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

  /** Lay each furniture slot out once (width = the content box) and remember
   *  each stack's height for the footer's bottom-edge placement. No grid
   *  context: Word snaps header/footer paragraphs to their natural line
   *  height — the body docGrid does not apply to the furniture story. */
  private layoutFurniture(): void {
    this.furnitureStacks = new Map();
    const f = this.ctx.furniture;
    if (!f) return;
    const measurer = new TextMeasurer(this.ctx.metrics);
    const lay = (blocks: typeof f.header): readonly LaidOutStackItem[] | undefined => {
      if (!blocks) return undefined;
      const laid = stackBlocks(blocks, this.ctx.flow.contentWidthPx, undefined, measurer);
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
      layer: "behind",
    };
    const { flow } = this.ctx;
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
    paintScene(tree, items, ctx);
    ctx.layer = "body";
    if (!this.storyEdit) {
      this.paintFurniture(tree, slot, ctx, flow);
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
      this.paintFurniture(tree, slot, ctx, flow);
      this.paintStoryBoundary(tree, ctx, flow);
    }
    // Render eagerly: Leafer's change-driven scheduling stalls when the App
    // was created while its view was offscreen (an IO callback during mount)
    // and never picks the page back up.
    app.forceRender();
  }

  /** Both furniture stacks for a page's slot, at their page positions. */
  private paintFurniture(
    tree: IGroup,
    slot: number,
    ctx: PaintContext,
    flow: ProjectedFlowBox,
  ): void {
    ctx.layer = "body";
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
