import {
  COMMAND_HOST_STYLE,
  TOOLTIP_PART,
  createIconSlot,
  forwardAttributes,
  preventFocusLoss,
  renderIcon,
  renderLabel,
} from "./command-helpers";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    ${COMMAND_HOST_STYLE}
    /* Labelled button (default = small): Fluent's button font is 14px
       (fontSizeBase300) and min-height 32px — too tall for a compact ribbon
       row. Drop the label to 12px and the button to 26px so a 3-row rb-grid of
       small commands stays tight. size="large" overrides both below; icon-only
       hides the label. size="small" is the same as the default, for callers
       that want an explicit, uniform size attribute across a group. */
    fluent-button {
      min-height: 26px;
    }
    .rb-label,
    :host([size="small"]) .rb-label {
      font-size: 12px;
    }
    /* size="large" — Office large button: icon stacked over label (column),
       larger glyph. Its height matches a large split (primary 56 + caret 14 =
       70) and its content is top-aligned (justify-content:flex-start), so the
       icon/label line up with a sibling large split's primary — the bottom gap
       stands in for the split's caret row. */
    :host([size="large"]) { flex-shrink: 0; }
    :host([size="large"]) fluent-button {
      flex-direction: column;
      justify-content: flex-start;
      min-width: 0;
      max-width: 90px;
      min-height: 70px;
      padding: 4px 12px;
    }
    :host([size="large"]) .rb-icon svg { width: 32px; height: 32px; }
    :host([size="large"]) .rb-label {
      font-size: 11px;
      text-align: center;
      line-height: 1.2;
      white-space: normal;
      word-break: break-word;
    }
  </style>
  <fluent-button id="target" part="button">
    <span class="rb-label"></span>
  </fluent-button>
  ${TOOLTIP_PART}`;

/**
 * `<docen-ribbon-button icon="cut" label="Cut" event="cut">` — a command that
 * wraps `<fluent-button appearance="subtle">` with a built-in `<fluent-tooltip>`
 * (label, or a `tooltip` attr override). The icon (a known Office name) goes
 * into `slot="start"`; with `icon-only` the label is hidden but still surfaces
 * as the tooltip. Click emits `command` with `{ event }`.
 */
class DocenRibbonButton extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["label", "icon", "event", "disabled", "icon-only", "tooltip"];
  }

  #btn?: HTMLElement;
  #icon?: HTMLSpanElement;
  #fwdCleanup?: () => void;
  #focusCleanup?: () => void;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "icon") this.#renderIcon();
    if (name === "label" || name === "tooltip") this.#renderLabel();
    if (name === "disabled") this.#reflect(name);
    if (name === "icon-only") {
      if (this.#icon) this.#icon.slot = this.hasAttribute("icon-only") ? "" : "start";
      this.#reflect(name);
      this.#renderLabel();
    }
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#btn = this.shadowRoot!.querySelector("fluent-button")!;
    this.#icon = createIconSlot(this.hasAttribute("icon-only") ? "" : "start");
    this.#btn.prepend(this.#icon);
    this.#renderIcon();
    this.#renderLabel();
    this.#reflect("disabled");
    this.#reflect("icon-only");
    // Default `subtle`; a caller can override via the `appearance` attribute.
    // `size` is docen's own (large = column layout, see CSS above), not fluent
    // — exclude it so it isn't forwarded onto the wrapped button.
    this.#fwdCleanup = forwardAttributes(this, this.#btn, { appearance: "subtle" }, ["size"]);
    // Keep the editor's selection on click — see preventFocusLoss.
    this.#focusCleanup = preventFocusLoss(this);
    this.#btn.addEventListener("click", () => this.#emit());
  }

  disconnectedCallback(): void {
    this.#fwdCleanup?.();
    this.#focusCleanup?.();
  }

  get event(): string {
    return this.getAttribute("event") ?? this.getAttribute("label") ?? "";
  }

  #tipText(): string {
    return this.getAttribute("tooltip") ?? this.getAttribute("label") ?? "";
  }

  #renderIcon(): void {
    if (this.#icon) renderIcon(this.#icon, this.getAttribute("icon") ?? "");
  }

  #renderLabel(): void {
    // icon-only: leave the visible label empty (it still feeds the tooltip).
    const visible = this.hasAttribute("icon-only") ? "" : (this.getAttribute("label") ?? "");
    renderLabel(this.shadowRoot!.querySelector(".rb-label")!, visible);
    renderLabel(this.shadowRoot!.querySelector(".rb-tip")!, this.#tipText());
  }

  #reflect(attr: string): void {
    if (!this.#btn) return;
    if (this.hasAttribute(attr)) this.#btn.setAttribute(attr, "");
    else this.#btn.removeAttribute(attr);
  }

  #emit(): void {
    if (this.hasAttribute("disabled")) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.event, source: this },
      }),
    );
  }
}

customElements.define("docen-ribbon-button", DocenRibbonButton);

export default DocenRibbonButton;
