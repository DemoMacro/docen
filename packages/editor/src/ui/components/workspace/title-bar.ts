import { FASTElement, css, customElement, html } from "@microsoft/fast-element";

const styles = css`
  :host {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 4px 8px;
    height: 40px;
    box-sizing: border-box;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    background: var(--docen-color-bg, #fff);
    font-family: "Segoe UI", "Segoe UI Web (West European)", system-ui, sans-serif;
    font-size: var(--docen-font-size-ribbon, 12px);
    color: var(--docen-color-text, #444);
  }
  [part="start"] {
    display: flex;
    align-items: center;
    gap: 4px;
    flex: 0 0 auto;
  }
  [part="search"] {
    flex: 1 1 auto;
    display: flex;
    justify-content: center;
  }
  [part="end"] {
    display: flex;
    align-items: center;
    gap: 4px;
    margin-inline-start: auto;
    flex: 0 0 auto;
  }
  /* Cap the centered search control so it doesn't span the whole header. */
  ::slotted([slot="search"]) {
    width: 100%;
    max-width: 480px;
  }
`;

const template = html<DocenTitleBar>`
  <div part="start"><slot name="start"></slot></div>
  <div part="search"><slot name="search"></slot></div>
  <div part="end"><slot name="end"></slot></div>
`;

/**
 * `<docen-title-bar>` — the workspace's global header (Office title bar):
 * `start` slot (logo / autosave / save / undo / redo / filename), `search`
 * slot (centered, capped width), `end` slot (user menu / window controls).
 * Layout only — editor packages slot their own controls; this component never
 * assumes specific content.
 */
@customElement({ name: "docen-title-bar", template, styles })
class DocenTitleBar extends FASTElement {}

export default DocenTitleBar;
