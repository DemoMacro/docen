import {
  COMMAND_HOST_STYLE,
  TOOLTIP_PART,
  createIconSlot,
  forwardAttributes,
  preventFocusLoss,
  renderIcon,
  renderLabel,
} from "./command-helpers";

export interface RibbonMenuItem {
  text: string;
  event?: string;
  value?: string;
  disabled?: boolean;
  /** Mutually-exclusive pick (Edit/View): renders a Fluent radio checkmark. */
  checked?: boolean;
}

// Per-instance CSS anchor name so each menu's popover aligns to its own
// trigger — without it, Fluent's ::slotted([popover]) default strands the
// dropdown at the viewport corner.
let seq = 0;

const template = document.createElement("template");
template.innerHTML = `
  <style>
    ${COMMAND_HOST_STYLE}
    /* Match <docen-ribbon-button>/<docen-ribbon-split-button>: Fluent's
       menu-button defaults to min-height 32px / 14px font — too tall for a
       compact ribbon row. Drop to 26px / 12px so a labelled menu sits flush
       with sibling ribbon commands. */
    fluent-menu-button {
      min-height: 26px;
    }
    .rb-label {
      font-size: 12px;
    }
  </style>
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
    return ["label", "icon", "items", "tooltip", "disabled"];
  }

  readonly #anchorId = `--rb-menu-${++seq}`;
  #trigger?: HTMLElement;
  #icon?: HTMLSpanElement;
  #fwdCleanup?: () => void;
  #focusCleanup?: () => void;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "icon") this.#renderIcon();
    if (name === "label" || name === "tooltip") this.#renderLabel();
    if (name === "items") this.#renderItems();
    if (name === "disabled") this.#reflectDisabled();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#trigger = this.shadowRoot!.querySelector("fluent-menu-button")!;
    // Anchor the popover list to the trigger (left-aligned under it). Without
    // this the dropdown opens at the viewport corner.
    const list = this.shadowRoot!.querySelector<HTMLElement>("fluent-menu-list")!;
    this.#trigger.style.anchorName = this.#anchorId;
    list.style.positionAnchor = this.#anchorId;
    list.style.insetInlineStart = "anchor(self-start)";
    list.style.insetInlineEnd = "unset";
    // The tooltip's connectedCallback runs before this host's, so its
    // anchor="target" resolves before anchorName is set and strands it at the
    // viewport corner — re-point it at this instance's anchor name.
    const tooltip = this.shadowRoot!.querySelector("fluent-tooltip");
    if (tooltip) (tooltip as HTMLElement).style.positionAnchor = this.#anchorId;
    this.#icon = createIconSlot();
    this.#trigger.prepend(this.#icon);
    this.#renderIcon();
    this.#renderLabel();
    this.#renderItems();
    this.#reflectDisabled();
    // Default `subtle`; a caller can override via the `appearance` attribute.
    this.#fwdCleanup = forwardAttributes(this, this.#trigger, { appearance: "subtle" });
    // Keep the editor focused on mousedown so opening the menu doesn't blur the
    // contenteditable — a blur/refocus race otherwise closes the popover right
    // after it opens (the "click several times before it appears" symptom).
    this.#focusCleanup = preventFocusLoss(this.#trigger);
  }

  disconnectedCallback(): void {
    this.#fwdCleanup?.();
    this.#focusCleanup?.();
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
      // A checked item is a mutually-exclusive mode pick (Edit/View) —
      // role="menuitemradio" + the `checked` attr renders Fluent's own
      // checkmark (no custom ::before). Plain items stay role="menuitem".
      if (item.checked) {
        menuItem.setAttribute("role", "menuitemradio");
        menuItem.setAttribute("checked", "");
      } else {
        menuItem.setAttribute("role", "menuitem");
      }
      menuItem.textContent = item.text;
      if (item.disabled) menuItem.setAttribute("disabled", "");
      menuItem.addEventListener("change", () => this.#emit(item));
      list.append(menuItem);
    }
  }

  /** Reflect the host `disabled` onto the menu trigger (mirrors
   *  <docen-ribbon-button>/<docen-ribbon-split-button>); #emit also bails when
   *  disabled, so a greyed menu fires nothing. */
  #reflectDisabled(): void {
    if (!this.#trigger) return;
    if (this.hasAttribute("disabled")) this.#trigger.setAttribute("disabled", "");
    else this.#trigger.removeAttribute("disabled");
  }

  #emit(item: RibbonMenuItem): void {
    if (this.hasAttribute("disabled")) return;
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
