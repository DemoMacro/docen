import {
  COMMAND_HOST_STYLE,
  TOOLTIP_PART,
  createIconSlot,
  forwardAttributes,
  renderIcon,
  renderLabel,
} from "./command-helpers";

export interface RibbonMenuItem {
  text: string;
  event?: string;
  value?: string;
  disabled?: boolean;
}

const template = document.createElement("template");
template.innerHTML = `
  <style>${COMMAND_HOST_STYLE}</style>
  <fluent-menu part="menu" style="--menu-max-height: auto;">
    <fluent-menu-button id="target" slot="trigger" part="button">
      <span slot="start" class="rb-icon"></span>
      <span class="rb-label"></span>
    </fluent-menu-button>
    <fluent-menu-list focusgroup="menu" part="list"></fluent-menu-list>
  </fluent-menu>
  ${TOOLTIP_PART}`;

/**
 * `<docen-ribbon-menu label="Calibri" items='[{...}]'>` — a command that opens a
 * native `<fluent-menu>` via a `<fluent-menu-button>` trigger (Fluent's own
 * caret). Items from the `items` JSON become `<fluent-menu-item role="menuitem">`
 * inside a `<fluent-menu-list focusgroup popover>`. Selecting one emits
 * `command` with `{ event, value }`. Fluent handles open/close, positioning,
 * focus and keyboard.
 */
class DocenRibbonMenu extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["label", "icon", "items", "tooltip"];
  }

  #trigger?: HTMLElement;
  #icon?: HTMLSpanElement;
  #fwdCleanup?: () => void;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "icon") this.#renderIcon();
    if (name === "label" || name === "tooltip") this.#renderLabel();
    if (name === "items") this.#renderItems();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#trigger = this.shadowRoot!.querySelector("fluent-menu-button")!;
    this.#icon = createIconSlot();
    this.#trigger.prepend(this.#icon);
    this.#renderIcon();
    this.#renderLabel();
    this.#renderItems();
    // Default `subtle`; a caller can override via the `appearance` attribute.
    this.#fwdCleanup = forwardAttributes(this, this.#trigger, { appearance: "subtle" });
  }

  disconnectedCallback(): void {
    this.#fwdCleanup?.();
  }

  get items(): RibbonMenuItem[] {
    try {
      return JSON.parse(this.getAttribute("items") ?? "[]") as RibbonMenuItem[];
    } catch {
      return [];
    }
  }

  #tipText(): string {
    return this.getAttribute("tooltip") ?? this.getAttribute("label") ?? "";
  }

  #renderIcon(): void {
    if (this.#icon) renderIcon(this.#icon, this.getAttribute("icon") ?? "");
  }

  #renderLabel(): void {
    renderLabel(this.shadowRoot!.querySelector(".rb-label")!, this.getAttribute("label") ?? "");
    renderLabel(this.shadowRoot!.querySelector(".rb-tip")!, this.#tipText());
  }

  #renderItems(): void {
    const list = this.shadowRoot?.querySelector("fluent-menu-list");
    if (!list) return;
    list.replaceChildren();
    for (const item of this.items) {
      const menuItem = document.createElement("fluent-menu-item");
      menuItem.setAttribute("role", "menuitem");
      menuItem.textContent = item.text;
      if (item.disabled) menuItem.setAttribute("disabled", "");
      menuItem.addEventListener("change", () => this.#emit(item));
      list.append(menuItem);
    }
  }

  #emit(item: RibbonMenuItem): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: {
          event: item.event ?? item.value ?? item.text,
          value: item.value,
          source: this,
        },
      }),
    );
  }
}

customElements.define("docen-ribbon-menu", DocenRibbonMenu);

export default DocenRibbonMenu;
