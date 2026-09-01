import { appendMenuItems } from "./ribbon/command-helpers";
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
  /** The document-level click listener that closes the menu (see #armDismiss). */
  #dismiss?: (event: Event) => void;

  disconnectedCallback(): void {
    this.#disarmDismiss();
  }

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
        this.#armDismiss();
      },
      true,
    );
    this.#renderItems();
  }

  /** Close on any click outside the menu list. Fluent's own light dismiss
   *  exempts the slotted trigger — which here wraps the WHOLE workspace — so
   *  a click on the document would never close the menu. Capture phase, so a
   *  stopPropagation deeper in the tree (canvas bridge, other menus) can't
   *  keep the stale menu up; a click INSIDE the list keeps it open. */
  #armDismiss(): void {
    this.#disarmDismiss();
    this.#dismiss = (event) => {
      if (!this.#list || event.composedPath().includes(this.#list)) return;
      this.#disarmDismiss();
      (this.#menu as unknown as { closeMenu: () => void }).closeMenu();
    };
    document.addEventListener("click", this.#dismiss, true);
  }

  #disarmDismiss(): void {
    if (this.#dismiss) document.removeEventListener("click", this.#dismiss, true);
    this.#dismiss = undefined;
  }

  get items(): RibbonMenuItem[] {
    try {
      return JSON.parse(this.getAttribute("items") ?? "[]") as RibbonMenuItem[];
    } catch {
      return [];
    }
  }

  #renderItems(): void {
    const list = this.shadowRoot?.querySelector<HTMLElement>("fluent-menu-list");
    if (list) appendMenuItems(list, this.items, (item) => this.#emit(item));
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
