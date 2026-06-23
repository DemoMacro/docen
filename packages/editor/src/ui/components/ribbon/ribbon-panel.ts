/**
 * `<docen-ribbon-panel value="home">…</docen-ribbon-panel>` — a transparent
 * pass-through container for one tab's groups. It owns neither layout nor
 * visibility: the parent `<docen-ribbon>` flags the active panel with
 * `data-active` and styles the slotted panels via `::slotted`. Keeping this
 * element free of any `:host{display:…}` means the parent's visibility rule
 * can never be shadowed.
 */
class DocenRibbonPanel extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["value"];
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      // DOM API over an innerHTML string: the shadow is just a single slot.
      this.attachShadow({ mode: "open" }).append(document.createElement("slot"));
    }
  }

  get value(): string {
    return this.getAttribute("value") ?? "";
  }
}

customElements.define("docen-ribbon-panel", DocenRibbonPanel);

export default DocenRibbonPanel;
