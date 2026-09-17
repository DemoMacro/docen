// The presentation element's static chrome: the shadow-root stylesheet and
// template. Mirrors the document chrome's shape — a <docen-workspace> shell
// with the thumbnails panel in the start pane and the slide surface in the
// default slot.

import { css, html, ref } from "@microsoft/fast-element";

/** Escape a host-supplied string for safe interpolation into innerHTML. The
 *  `filename` attribute comes from a user-selected File.name, which can
 *  contain markup — without escaping it flows into #renderHeader's template
 *  and executes. */
export const escapeHtml = (s: string): string =>
  s.replace(/[&<>"']/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&#39;",
  );

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
  /* The drawing gridlines: a non-interactive checker over the whole slide
     strip, sized per-zoom from the host (PowerPoint's 0.5" grid). */
  .gridlines {
    position: absolute;
    inset: 0;
    z-index: 2;
    pointer-events: none;
    background-image:
      linear-gradient(to right, rgba(102, 102, 102, 0.28) 1px, transparent 1px),
      linear-gradient(to bottom, rgba(102, 102, 102, 0.28) 1px, transparent 1px);
  }
  .gridlines[hidden] {
    display: none;
  }
  /* The in-place shape text editor: a bare textarea floating over the shape
     (same layer as the selection overlay), framed like it so the edit reads
     as belonging to the shape. The projection member supplies the real
     insets/face/color per shape — these are the bare-box fallbacks. */
  .shape-text-editor {
    position: absolute;
    z-index: 6;
    box-sizing: border-box;
    margin: 0;
    padding: 4.8px 9.6px;
    background: rgba(255, 255, 255, 0.92);
    border: 1.5px solid #2b7cd3;
    outline: none;
    resize: none;
    overflow: hidden;
    font-family: inherit;
    line-height: normal;
    white-space: pre-wrap;
  }
`;

/** The element surface the template binds to — the ref targets only. The
 *  element class structurally matches it (the observables carry the nodes). */
interface PresentationTemplateRefs extends HTMLElement {
  thumbStrip?: HTMLElement;
  thumbSelection?: HTMLElement;
  gridlines?: HTMLElement;
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
        <div class="gridlines" ${ref("gridlines")} hidden></div>
      </div>
    </docen-document-area>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <input type="file" id="file-input" accept=".pptx" hidden />
  <input type="file" id="picture-input" accept="image/png,image/jpeg,image/gif,image/bmp" hidden />
`;
