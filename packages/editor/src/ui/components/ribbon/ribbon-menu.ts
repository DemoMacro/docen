import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { COMMAND_HOST_STYLE, renderIcon } from "./command-helpers";

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

const styles = css`
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
`;

const template = html<DocenRibbonMenu>`
  <fluent-menu part="menu" style="--menu-max-height: auto;">
    <fluent-menu-button
      id="target"
      slot="trigger"
      part="button"
      appearance="${(x) => x.appearance ?? "subtle"}"
      ?disabled="${(x) => x.disabled}"
      ${ref("trigger")}
    >
      <span slot="start" class="rb-icon" ${ref("iconSlot")}></span>
      <span class="rb-label">${(x) => x.label}</span>
    </fluent-menu-button>
    <fluent-menu-list focusgroup="menu" part="list" ${ref("list")}></fluent-menu-list>
  </fluent-menu>
  <fluent-tooltip anchor="target" positioning="top" ${ref("tooltipEl")}>
    <span class="rb-tip">${(x) => x.tooltipText}</span>
  </fluent-tooltip>
`;

/**
 * `<docen-ribbon-menu label="Calibri" items='[{...}]'>` — a command that opens a
 * native `<fluent-menu>` via a `<fluent-menu-button>` trigger (Fluent's own
 * caret). Items from the `items` JSON become `<fluent-menu-item role="menuitem">`
 * inside a `<fluent-menu-list focusgroup popover>`. Selecting one emits
 * `command` with `{ event, value }`. Fluent handles open/close, positioning,
 * focus and keyboard.
 */
@customElement({ name: "docen-ribbon-menu", template, styles })
class DocenRibbonMenu extends FASTElement {
  @attr label?: string;
  @attr icon?: string;
  @attr tooltip?: string;
  @attr appearance?: string;
  @attr({ mode: "boolean" }) disabled?: boolean;
  @attr items?: string;

  @observable trigger?: HTMLElement;
  @observable list?: HTMLElement;
  @observable tooltipEl?: HTMLElement;
  @observable iconSlot?: HTMLSpanElement;

  readonly anchorId = `--rb-menu-${++seq}`;

  get tooltipText(): string {
    return this.tooltip || this.label || "";
  }
  get parsedItems(): RibbonMenuItem[] {
    try {
      return JSON.parse(this.items ?? "[]") as RibbonMenuItem[];
    } catch {
      return [];
    }
  }

  iconChanged(): void {
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
  }

  itemsChanged(): void {
    this.renderItems();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.applyAnchor();
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
    this.renderItems();
    // Keep the editor focused on mousedown so opening the menu doesn't blur the
    // contenteditable — a blur/refocus race otherwise closes the popover right
    // after it opens (the "click several times before it appears" symptom).
    this.trigger?.addEventListener("mousedown", this.onMousedown, { capture: true });
  }

  disconnectedCallback(): void {
    this.trigger?.removeEventListener("mousedown", this.onMousedown, { capture: true });
    super.disconnectedCallback();
  }

  private readonly onMousedown = (event: Event): void => event.preventDefault();

  private applyAnchor(): void {
    // Anchor the popover list to the trigger (left-aligned under it). Without
    // this the dropdown opens at the viewport corner.
    if (this.trigger) this.trigger.style.anchorName = this.anchorId;
    if (this.list) {
      this.list.style.positionAnchor = this.anchorId;
      this.list.style.insetInlineStart = "anchor(self-start)";
      this.list.style.insetInlineEnd = "unset";
    }
    // The tooltip's connectedCallback runs before this host's, so its anchor
    // resolves before anchorName is set — re-point it at this instance's anchor.
    if (this.tooltipEl) this.tooltipEl.style.positionAnchor = this.anchorId;
  }

  private renderItems(): void {
    if (!this.list) return;
    this.list.replaceChildren();
    for (const item of this.parsedItems) {
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
      menuItem.addEventListener("change", () => this.emit(item));
      this.list.append(menuItem);
    }
  }

  private emit(item: RibbonMenuItem): void {
    if (this.disabled) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: item.event ?? item.value ?? item.text, value: item.value, source: this },
      }),
    );
  }
}

export default DocenRibbonMenu;
