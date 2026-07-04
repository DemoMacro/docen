import { FASTElement, attr, customElement, html } from "@microsoft/fast-element";

/**
 * `<docen-ribbon-panel value="home">…</docen-ribbon-panel>` — a transparent
 * pass-through container for one tab's groups. It owns neither layout nor
 * visibility: the parent `<docen-ribbon>` flags the active panel with
 * `data-active` and styles the slotted panels via `::slotted`. Keeping this
 * element free of any `:host{display:…}` means the parent's visibility rule
 * can never be shadowed.
 */
const template = html<DocenRibbonPanel>`<slot></slot>`;

@customElement({ name: "docen-ribbon-panel", template })
class DocenRibbonPanel extends FASTElement {
  @attr value?: string;
}

export default DocenRibbonPanel;
