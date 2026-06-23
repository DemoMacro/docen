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
  <style>${COMMAND_HOST_STYLE}</style>
  <fluent-toggle-button id="target" part="button">
    <span class="rb-label"></span>
  </fluent-toggle-button>
  ${TOOLTIP_PART}`;

/**
 * `<docen-ribbon-toggle-button icon="bold" label="Bold" event="bold" icon-only>`
 * — a two-state command (Office "Toggle Button") wrapping
 * `<fluent-toggle-button appearance="subtle">` with a built-in tooltip. Fluent
 * owns the pressed affordance; `pressed` (attr) seeds the initial state and the
 * getter reads Fluent's live value. Click emits `command` with
 * `{ event, pressed }`.
 */
class DocenRibbonToggleButton extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["label", "icon", "event", "pressed", "disabled", "icon-only", "tooltip"];
  }

  #btn?: HTMLElement;
  #icon?: HTMLSpanElement;
  #fwdCleanup?: () => void;
  #focusCleanup?: () => void;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "icon") this.#renderIcon();
    if (name === "label" || name === "tooltip") this.#renderLabel();
    if (name === "pressed") this.#syncPressed();
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
    this.#btn = this.shadowRoot!.querySelector("fluent-toggle-button")!;
    this.#icon = createIconSlot(this.hasAttribute("icon-only") ? "" : "start");
    this.#btn.prepend(this.#icon);
    this.#renderIcon();
    this.#renderLabel();
    this.#syncPressed();
    this.#reflect("disabled");
    this.#reflect("icon-only");
    // Default `subtle`; a caller can override via the `appearance` attribute.
    this.#fwdCleanup = forwardAttributes(this, this.#btn, { appearance: "subtle" });
    // Keep the editor's selection on click — see preventFocusLoss.
    this.#focusCleanup = preventFocusLoss(this);
    // Defer until Fluent has toggled its internal pressed state for this click.
    this.#btn.addEventListener("click", () => queueMicrotask(() => this.#emit()));
  }

  disconnectedCallback(): void {
    this.#fwdCleanup?.();
    this.#focusCleanup?.();
  }

  get event(): string {
    return this.getAttribute("event") ?? this.getAttribute("label") ?? "";
  }

  get pressed(): boolean {
    return (this.#btn as { pressed?: boolean } | undefined)?.pressed ?? false;
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

  #syncPressed(): void {
    if (!this.#btn) return;
    // `pressed` is a boolean attribute on fluent-toggle-button; reflecting it
    // seeds the observable, and the getter above reads the live value.
    if (this.hasAttribute("pressed")) this.#btn.setAttribute("pressed", "");
    else this.#btn.removeAttribute("pressed");
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
        detail: { event: this.event, pressed: this.pressed, source: this },
      }),
    );
  }
}

customElements.define("docen-ribbon-toggle-button", DocenRibbonToggleButton);

export default DocenRibbonToggleButton;
