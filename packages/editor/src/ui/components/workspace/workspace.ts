import { FASTElement, css, customElement, html } from "@microsoft/fast-element";

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    min-height: 0;
    height: 100%;
    background: var(--docen-color-bg, #fff);
    font-family:
      "Segoe UI",
      system-ui,
      -apple-system,
      sans-serif;
    color: var(--docen-color-text, #3b3b3b);
  }
  .rb-shell-header {
    flex: 0 0 auto;
  }
  /* position: relative gives the auto-hide ribbon host (position: absolute)
     its positioning context so it overlays in place when revealed. */
  .rb-shell-ribbon {
    flex: 0 0 auto;
    position: relative;
  }
  .rb-shell-content {
    flex: 1 1 auto;
    min-height: 0;
    display: flex;
  }
  .rb-shell-content .rb-canvas {
    flex: 1 1 auto;
    min-width: 0;
    display: flex;
  }
  .rb-shell-content .rb-aside {
    flex: 0 0 auto;
    min-width: 0;
    overflow: hidden;
  }
  .rb-shell-status {
    flex: 0 0 auto;
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 4px 12px;
    min-height: 28px;
    box-sizing: border-box;
    background: var(--docen-color-status-bg, #f0f0f0);
    border-top: 1px solid var(--docen-color-divider, #e1e1e1);
    font-size: var(--docen-font-size-ribbon, 12px);
    color: var(--docen-color-text-muted, #6b6b6b);
  }
  /* Full Screen (ribbon auto-hide): hide the status bar — the app header and
     side panes stay visible per Office behavior. */
  :host([data-fullscreen]) .rb-shell-status {
    display: none;
  }
  /* Print: hide everything but the canvas (the printed page). */
  @media print {
    :host {
      display: block;
      height: auto;
      background: #fff;
    }
    .rb-shell-header,
    .rb-shell-ribbon,
    .rb-shell-status,
    .rb-aside {
      display: none !important;
    }
    .rb-shell-content {
      display: block;
    }
    .rb-canvas {
      display: block;
    }
  }
`;

const template = html<DocenWorkspace>`
  <div class="rb-shell-header" part="header"><slot name="header"></slot></div>
  <div class="rb-shell-ribbon" part="ribbon"><slot name="ribbon"></slot></div>
  <div class="rb-shell-content" part="content">
    <aside class="rb-aside" part="task-pane-start"><slot name="task-pane-start"></slot></aside>
    <div class="rb-canvas" part="canvas"><slot></slot></div>
    <aside class="rb-aside" part="task-pane-end"><slot name="task-pane-end"></slot></aside>
  </div>
  <div class="rb-shell-status" part="status"><slot name="status"></slot></div>
`;

/**
 * `<docen-workspace>` — the editor application shell: global header (very
 * top) + Ribbon + canvas (middle, default slot) flanked by Task Panes
 * (start/end) + status bar (bottom). Editor packages slot their engine
 * (Tiptap/LeaferJS/RevoGrid) into the default slot and Office-style side
 * panels into `task-pane-*`.
 *
 * Also the locale provider for the Fluent components: set `lang` to override
 * the page locale for component-internal strings (see src/i18n/localize.ts).
 */
@customElement({ name: "docen-workspace", template, styles })
class DocenWorkspace extends FASTElement {}

export default DocenWorkspace;
