import type { RibbonMenuItem } from "./ribbon/ribbon-menu";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host { display: flex; flex-direction: column; }
    /* The trigger wraps the slotted workspace and fills the host so it stays a
       valid anchor/ARIA owner for fluent-menu. */
    [part="trigger"] { flex: 1; min-height: 0; display: flex; }
  </style>
  <fluent-menu part="menu" open-on-context style="--menu-max-height: auto;">
    <div part="trigger" slot="trigger"><slot></slot></div>
    <fluent-menu-list focusgroup="menu" part="list"></fluent-menu-list>
  </fluent-menu>`;

/**
 * `<docen-context-menu items='[{...}]'>…editor content…</docen-context-menu>` —
 * wraps `<fluent-menu>` so right-clicking the slotted workspace opens a Fluent
 * menu at the cursor. fluent-menu's built-in `open-on-context` ignores the
 * cursor (it anchors the popover to the trigger's top-left), so this component
 * listens for `contextmenu` itself, pins the menu-list to the cursor, then
 * calls the menu's `openMenu()`. Items become `<fluent-menu-item>`s; selecting
 * one emits `command` with `{ event, value }`. Fluent owns focus and keyboard.
 */
class DocenContextMenu extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["items"];
  }

  attributeChangedCallback(name: string): void {
    if (name === "items") this.#renderItems();
  }

  #menu?: HTMLElement;
  #list?: HTMLElement;

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#menu = this.shadowRoot!.querySelector("fluent-menu")!;
    this.#list = this.shadowRoot!.querySelector("fluent-menu-list")!;
    // The whole workspace is the right-click target. We open the menu ourselves
    // and pin the list to the cursor. open-on-context is set on the menu only so
    // fluent takes its contextmenu branch (and does NOT add the default
    // trigger-click -> toggleMenu listener, which would open on a left click).
    // We capture + stopPropagation so the menu's own contextmenu handler doesn't
    // also fire and fight our cursor positioning.
    this.addEventListener(
      "contextmenu",
      (event) => {
        event.preventDefault();
        if (!this.#list || !this.#menu) return;
        this.#list.style.top = `${event.clientY}px`;
        this.#list.style.left = `${event.clientX}px`;
        event.stopPropagation();
        (this.#menu as unknown as { openMenu: () => void }).openMenu();
      },
      true,
    );
    this.#renderItems();
  }

  get items(): RibbonMenuItem[] {
    try {
      return JSON.parse(this.getAttribute("items") ?? "[]") as RibbonMenuItem[];
    } catch {
      return [];
    }
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

customElements.define("docen-context-menu", DocenContextMenu);

export default DocenContextMenu;
