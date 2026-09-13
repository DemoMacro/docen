/**
 * `<docen-presentation>` — PPTX viewer element.
 *
 * The minimal open-and-render loop: parse the .pptx through @docen/pptx,
 * project the slides into drawing members, and paint them with the shared
 * core painter on a Leafer stage — slides stacked vertically like document
 * pages. Editing (selection, gestures, ribbon) rides on this base in later
 * batches.
 */

import { paintMembers } from "@docen/core";
import { browserFontMetrics } from "@docen/layout";
import { parsePresentation, projectPresentation, type ProjectedPresentation } from "@docen/pptx";
import type { DataType } from "@office-open/core";
import { App, Group, Rect, type IGroup } from "leafer-ui";

/** Gap between consecutive slides, px. */
const SLIDE_GAP_PX = 24;

class DocenPresentation extends HTMLElement {
  #pres: ProjectedPresentation | null = null;
  #app: App | null = null;
  #scroller: HTMLDivElement | null = null;
  #stage: HTMLDivElement | null = null;

  connectedCallback(): void {
    if (this.shadowRoot) return;
    const root = this.attachShadow({ mode: "open" });
    const style = document.createElement("style");
    style.textContent = `
      :host { display: flex; height: 100%; }
      .scroller {
        flex: 1;
        overflow: auto;
        background: #e6e6e6;
      }
      .stage { margin: 0 auto; }
      .stage canvas { display: block; }
    `;
    root.append(style);
    this.#scroller = document.createElement("div");
    this.#scroller.className = "scroller";
    this.#stage = document.createElement("div");
    this.#stage.className = "stage";
    this.#scroller.append(this.#stage);
    root.append(this.#scroller);
    if (this.#pres) this.#render();
  }

  disconnectedCallback(): void {
    this.#app?.destroy();
    this.#app = null;
  }

  /** Parse a .pptx payload and render every slide. */
  async openPresentation(data: DataType): Promise<void> {
    const pres = await parsePresentation(data);
    this.#pres = projectPresentation(pres);
    if (this.shadowRoot) this.#render();
  }

  #render(): void {
    const pres = this.#pres;
    const stage = this.#stage;
    if (!pres || !stage) return;
    // Fresh tree per open — slides are static until the editing engine lands.
    this.#app?.destroy();
    stage.replaceChildren();
    // Slide strip: one gap above, between, and below the slides.
    const stripHeight = pres.slides.length * (pres.heightPx + SLIDE_GAP_PX) + SLIDE_GAP_PX;
    stage.style.width = `${pres.widthPx}px`;
    stage.style.height = `${stripHeight}px`;
    const app = new App({
      view: stage,
      fill: "transparent",
      tree: { type: "design" },
      move: { disabled: true },
      wheel: { disabled: true },
    });
    this.#app = app;
    const tree = app.tree as unknown as IGroup;
    let y = SLIDE_GAP_PX;
    for (let i = 0; i < pres.slides.length; i++) {
      const slide = pres.slides[i]!;
      const slideGroup = new Group({ x: 0, y });
      slideGroup.add(
        new Rect({
          width: pres.widthPx,
          height: pres.heightPx,
          fill: slide.background ? `#${slide.background}` : "#ffffff",
          stroke: "#c4c4c4",
          strokeWidth: 1,
        }),
      );
      paintMembers(slideGroup, slide.members, 0, 0, {
        metrics: browserFontMetrics,
        flow: {
          pageWidthPx: pres.widthPx,
          pageHeightPx: pres.heightPx,
          contentWidthPx: pres.widthPx,
          contentHeightPx: pres.heightPx,
          contentLeftPx: 0,
          contentTopPx: 0,
        },
        pageIndex: i,
        pageCount: pres.slides.length,
        layer: "body",
        rerender: () => app.forceRender(),
      });
      tree.add(slideGroup);
      y += pres.heightPx + SLIDE_GAP_PX;
    }
  }
}

customElements.define("docen-presentation", DocenPresentation);

export default DocenPresentation;
