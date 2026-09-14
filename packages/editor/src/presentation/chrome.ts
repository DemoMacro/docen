// The presentation element's static chrome: the shadow-root stylesheet and
// template. Mirrors the document chrome's shape — a <docen-workspace> shell
// with the thumbnails panel in the start pane and the slide surface in the
// default slot.

import { css, html, ref } from "@microsoft/fast-element";

export { escapeHtml } from "../document/chrome";

export const presentationStyles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
  }
  /* The thumbnails rail (PowerPoint's slide panel): a fixed column of small
     deck previews; hidden until a deck opens. The selection frame is an
     absolutely positioned sibling of the thumbnail canvas — the canvas is
     one Leafer surface for the whole deck, so the frame moves by transform. */
  .slides-panel {
    width: 188px;
    overflow-y: auto;
    overscroll-behavior: contain;
    background: var(--docen-color-bg, #fff);
    border-inline-end: 1px solid var(--docen-color-divider, #e1e1e1);
  }
  .slides-panel[hidden] {
    display: none;
  }
  .thumb-strip {
    position: relative;
    margin: 12px auto;
    width: fit-content;
  }
  .thumb-stage canvas {
    display: block;
    cursor: pointer;
  }
  .thumb-selection {
    position: absolute;
    inset-inline: 0;
    top: 0;
    height: 0;
    border: 2px solid var(--docen-color-accent, #0f6cbd);
    pointer-events: none;
    box-sizing: border-box;
  }
  /* The slide surface — document-area is the scroll container; this wrapper
     centers the Leafer stage like the document's page column. It anchors the
     selection overlay (position:relative), so the slide-top gap is a margin:
     an absolutely positioned child measures from the padding box edge. */
  .docen-canvas {
    position: relative;
    width: fit-content;
    margin: 32px auto 0;
  }
  .stage canvas {
    display: block;
  }
`;

/** The element surface the template binds to — the ref targets only. The
 *  element class structurally matches it (the observables carry the nodes). */
interface PresentationTemplateRefs extends HTMLElement {
  thumbStrip?: HTMLElement;
  thumbSelection?: HTMLElement;
}

export const presentationTemplate = html<PresentationTemplateRefs>`
  <docen-workspace>
    <docen-title-bar slot="header" part="header"></docen-title-bar>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <div class="slides-panel" slot="task-pane-start" part="slides-panel" hidden>
      <div class="thumb-strip" ${ref("thumbStrip")}>
        <div class="thumb-stage"></div>
        <div class="thumb-selection" ${ref("thumbSelection")}></div>
      </div>
    </div>
    <docen-document-area>
      <div class="docen-canvas" part="page">
        <div class="stage"></div>
      </div>
    </docen-document-area>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <input type="file" id="file-input" accept=".pptx" hidden />
  <input type="file" id="picture-input" accept="image/png,image/jpeg,image/gif,image/bmp" hidden />
`;
