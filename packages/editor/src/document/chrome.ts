// The element's static chrome: the shadow-root stylesheet and template, plus
// the HTML escaping the filename header interpolation needs.

import { css, html } from "@microsoft/fast-element";

/** Escape a host-supplied string for safe interpolation into innerHTML. The
 *  `filename` attribute comes from a user-selected File.name at openDOCX, which
 *  can contain markup — without escaping it flows into #renderHeader's template
 *  and executes. */
export const escapeHtml = (s: string): string =>
  s.replace(/[&<>"']/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&#39;",
  );

export const documentStyles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    /* Anchors the input layer (see the template comment there). */
    position: relative;
  }
  .input-layer {
    position: absolute;
    inset: 0;
    /* The layer itself must never intercept pointer input — only the bridge's
       programmatic textarea focus uses it. */
    pointer-events: none;
  }
  /* Office ribbon group layout helpers — a large button beside stacked rows of
       small icon-only buttons. Applied to light-DOM wrappers in the ribbon. */
  .rb-col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .rb-row {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 2px;
    flex-wrap: wrap;
  }
  /* Small icon-only buttons as a 3-row column-flow grid: buttons stack into
       columns of ≤3 (Word's compact group layout), not a flat single row. */
  .rb-grid {
    display: grid;
    grid-template-rows: repeat(3, auto);
    grid-auto-flow: column;
    gap: 2px;
    align-content: start;
  }
  .rb-vsep {
    width: 1px;
    align-self: stretch;
    background: var(--docen-color-divider, #e1e1e1);
    margin: 0 2px;
  }
  .avatar {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    background: var(--docen-color-brand, #0078d4);
    color: #fff;
    font-size: 10px;
    font-weight: 600;
    margin-inline-end: 4px;
  }
  .avatar-img {
    object-fit: cover;
    background: none;
  }
  /* The canvas surface — the scroll container sits one level up (the
       document-area); this wrapper just anchors the edit bridge's textarea and
       caret overlays (position:relative). cursor:text is the editing surface's
       I-beam, like Word's page area. */
  .docen-canvas {
    position: relative;
    width: fit-content;
    margin: 0 auto;
    padding: 32px 0;
    cursor: text;
  }
  /* Open-progress veil over the canvas (Word centers its opening spinner in
       the document area too): label + Fluent progress bar, centered on a
       translucent white wash so the not-yet-laid-out document doesn't flash
       behind it. The area is the scroll container, so the veil is its FIRST
       child, sticky at top, and one area-height tall (= the visible region,
       so the center lands mid-viewport). While shown, the veil adds its
       height to the scroll range — so freeze the scroller for the duration
       (nothing behind it is worth scrolling to). */
  docen-document-area:has(> .load-veil:not([hidden])) {
    overflow: hidden;
  }
  .load-veil {
    position: sticky;
    top: 0;
    /* 100% = the area's content box (its 24px paddings stay uncovered — under
       the translucent wash the not-yet-replaced document shows as a hairline
       edge, invisible against a blank first load). */
    height: 100%;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 12px;
    background: rgba(255, 255, 255, 0.82);
    z-index: 40;
  }
  .load-veil[hidden] {
    display: none;
  }
  .load-veil fluent-progress-bar {
    width: 240px;
  }
  .load-veil .load-label {
    font-size: 13px;
    color: var(--docen-color-text-2, #424242);
  }
  /* Grey the "Auto-save" label to match its disabled switch (skeleton
       feature), so the label + switch read as one unavailable control, like
       ribbon skeleton buttons. Lifts automatically once the switch loses
       disabled. */
  .autosave-label:has(+ fluent-switch[disabled]) {
    color: var(--docen-color-text-3, #8a8a8a);
  }
  /* Find Results — Office-style match list: each hit rendered with surrounding
       context and a data-from/to for click-to-jump. Padding keeps items off the
       pane edge (the previous "N matches" text butted right against it). */
  .search-results {
    padding: 6px 8px;
    box-sizing: border-box;
  }
  .search-results .result-count {
    font-size: 12px;
    color: var(--docen-color-marks, #6e6e6e);
    padding: 2px 4px 8px;
  }
  .search-results .result-item {
    display: block;
    width: 100%;
    text-align: start;
    border: none;
    background: transparent;
    padding: 5px 8px;
    margin-block-end: 2px;
    border-radius: 4px;
    font-family: inherit;
    font-size: 12px;
    line-height: 1.45;
    color: #3b3b3b;
    cursor: pointer;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .search-results .result-item:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .search-results .result-item mark {
    background: rgba(255, 235, 59, 0.85);
    color: inherit;
    font-weight: 600;
  }
`;

export const documentTemplate = html`
  <docen-workspace>
    <docen-title-bar slot="header" part="header"></docen-title-bar>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <docen-task-pane slot="task-pane-start" position="start" part="nav-pane">
      <docen-navigation-pane>
        <docen-outline slot="headings"></docen-outline>
        <div class="search-results" slot="results" part="search-results"></div>
      </docen-navigation-pane>
    </docen-task-pane>
    <docen-document-area>
      <div class="load-veil" part="load-veil" hidden>
        <fluent-progress-bar></fluent-progress-bar>
        <span class="load-label"></span>
      </div>
      <docen-context-menu part="context-menu">
        <div class="docen-canvas" part="page"></div>
      </docen-context-menu>
    </docen-document-area>
    <docen-task-pane slot="task-pane-end" position="end" part="props-pane">
      <slot name="properties">
        <docen-format-pane></docen-format-pane>
      </slot>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="comments-pane" title="Comments">
      <docen-comments-pane></docen-comments-pane>
    </docen-task-pane>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <!-- The edit bridge's textarea lives here, at the shadow root: inside the
       workspace it would sit under docen-context-menu, whose fluent-menu
       treats Space/Enter as menu keys and preventDefaults them — killing the
       textarea's beforeinput (spaces and Enter silently dropped). -->
  <div class="input-layer" part="input-layer"></div>
  <docen-options-dialog part="options"></docen-options-dialog>
  <docen-word-count-dialog part="word-count"></docen-word-count-dialog>
  <docen-symbol-dialog part="symbol"></docen-symbol-dialog>
  <docen-paragraph-dialog part="paragraph"></docen-paragraph-dialog>
  <docen-page-setup-dialog part="page-setup"></docen-page-setup-dialog>
  <docen-table-dialog part="table"></docen-table-dialog>
  <docen-columns-dialog part="columns"></docen-columns-dialog>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <input type="file" id="file-input" accept=".docx,.md,.markdown" hidden />
  <input type="file" id="image-input" accept="image/*" hidden />
`;
