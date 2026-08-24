import type { ProjectedFlowBox } from "@docen/docx/layout";
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
import type { FlowPage, FontMetrics } from "@docen/layout";
import { App } from "leafer-ui";

import { paintScene } from "./scene";

const PAGE_GAP = 24;

/** Font metrics (baseline ratios for half-leading) + page geometry. */
export interface CanvasStageContext {
  metrics: FontMetrics;
  flow: ProjectedFlowBox;
}

export class CanvasStage {
  readonly shell: HTMLElement;
  private readonly slots: { el: HTMLElement; app: App | null }[] = [];
  private readonly io: IntersectionObserver;
  private pages: FlowPage[] = [];

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
  }

  /** Lay out page slots for a flow result and repaint visible pages. */
  sync(pages: FlowPage[], flow: ProjectedFlowBox): void {
    this.pages = pages;
    this.ctx.flow = flow;
    const w = Math.ceil(flow.pageWidthPx);
    const h = Math.ceil(flow.pageHeightPx);

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
    for (const slot of this.slots) slot.app?.destroy();
    this.shell.remove();
  }

  private ensure(slot: { el: HTMLElement; app: App | null }): void {
    if (slot.app) return;
    const app = new App({
      view: slot.el,
      fill: "transparent",
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

  private repaint(app: App, index: number): void {
    // app.tree is an ILeafer (extends IGroup) — clear/add come with it.
    const tree = app.tree;
    if (!tree) return;
    tree.clear();
    paintScene(tree, this.pages[index]?.items ?? [], this.ctx);
    // Render eagerly: Leafer's change-driven scheduling stalls when the App
    // was created while its view was offscreen (an IO callback during mount)
    // and never picks the page back up.
    app.forceRender();
  }
}
