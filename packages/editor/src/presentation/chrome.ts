// The presentation element's static chrome: the shadow-root stylesheet and
// template. Mirrors the document chrome's shape — a <docen-workspace> shell
// with the slide surface in the default slot — trimmed to the surfaces the
// presentation editor fills today (no task panes or dialogs yet).

import { css, html } from "@microsoft/fast-element";

export { escapeHtml } from "../document/chrome";

export const presentationStyles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
  }
  /* The slide surface — document-area is the scroll container; this wrapper
     centers the Leafer stage like the document's page column. No cursor
     styling: there is no text caret to point at yet. */
  .docen-canvas {
    width: fit-content;
    margin: 0 auto;
    padding: 32px 0;
  }
  .stage canvas {
    display: block;
  }
`;

export const presentationTemplate = html`
  <docen-workspace>
    <docen-title-bar slot="header" part="header"></docen-title-bar>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <docen-document-area>
      <div class="docen-canvas" part="page">
        <div class="stage"></div>
      </div>
    </docen-document-area>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <input type="file" id="file-input" accept=".pptx" hidden />
`;
