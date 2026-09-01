import { FASTElement, css, customElement, html } from "@microsoft/fast-element";

const styles = css`
  :host {
    display: block;
    flex: 1 1 auto;
    min-width: 0;
    min-height: 0;
    overflow: auto;
    background: var(--docen-color-canvas, #ffffff);
    padding: var(--docen-page-gap, 24px);
  }
  /* Print: drop the scroll chrome so only the pages print. */
  @media print {
    :host {
      overflow: visible;
      padding: 0;
      background: #fff;
    }
  }
`;

const template = html<DocenDocumentArea>`<slot></slot>`;

/**
 * `<docen-document-area>` — the editor workspace surface: a scrolling
 * container. The editor slots the context menu wrapping the canvas host
 * (`.docen-canvas`, the Leafer stage's mount point) into the default slot;
 * page geometry lives entirely in the canvas stage, which sizes each page
 * slot from the layout flow — this element only owns the scroll chrome.
 */
@customElement({ name: "docen-document-area", template, styles })
class DocenDocumentArea extends FASTElement {}

export default DocenDocumentArea;
