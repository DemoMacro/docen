import {
  COMMAND_HOST_STYLE,
  TOOLTIP_PART,
  createIconSlot,
  forwardAttributes,
  preventFocusLoss,
  renderIcon,
  renderLabel,
} from "./command-helpers";
import type { RibbonMenuItem } from "./ribbon-menu";

// Per-instance CSS anchor name so each split's dropdown aligns to its own
// button (left edge), not Fluent's default caret right-edge — which on a
// narrow icon-only split pushes the dropdown far to the left.
let seq = 0;

const template = document.createElement("template");
template.innerHTML = `
  <style>
    ${COMMAND_HOST_STYLE}
    /* Labelled split primary (default = small): Fluent's button font is 14px
       and min-height 32px — too tall for a compact ribbon row. Drop the label
       to 12px and the primary to 26px so the split stays tight. size="large"
       overrides both below; icon-only hides the label. size="small" is the same
       as the default, for an explicit, uniform size attribute. */
    fluent-button[slot="primary-action"] {
      min-height: 26px;
    }
    .rb-label,
    :host([size="small"]) .rb-label {
      font-size: 12px;
    }
    /* size="large" — Office large split button: the primary stacks icon over
       label (column), the caret drops below it (fluent-menu data-vertical).
       The primary aligns its content to the top (justify-content:flex-start)
       so its icon/label line up with a sibling <docen-ribbon-button
       size="large">, whose matching bottom gap stands in for the caret row. */
    :host([size="large"]) { flex-shrink: 0; }
    :host([size="large"]) fluent-button {
      flex-direction: column;
      justify-content: flex-start;
      min-width: 0;
      max-width: 90px;
      min-height: 56px;
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
    /* The caret is an icon-only fluent-menu-button: its :host([icon-only])
       clamps width to 32px and :host keeps 12px inline padding, so the 14px
       glyph is pushed off-center (horizontal caret looks right-shifted) and
       the bar won't span the primary (vertical caret looks too short, making
       the large split lopsided). Drop the inline padding everywhere and
       stretch full-width in the large column. Specificity is raised past the
       Fluent :host / :host([icon-only]) rules (shadow :host loses ties to a
       higher-specificity outer rule). */
    :host [part="caret"] {
      padding-inline: 0;
    }
    /* Non-large split (icon-only): a narrow caret hugging the primary — the
       Office compact split. Fluent's icon-only button clamps max-width, so lift
       it and pin a tight width; the chevron centers within. */
    :host(:not([size="large"])) [part="caret"] {
      width: 14px;
      min-width: 14px;
      max-width: none;
    }
    :host([size="large"]) [part="caret"] {
      width: 100%;
      min-width: 0;
      max-width: none;
      /* Office large split: the caret is a short bar under the primary, not a
         full-height button — menu-button defaults to min-height:32px. */
      min-height: 14px;
    }
  </style>
  <fluent-menu split part="menu" style="--menu-max-height: auto;">
    <fluent-button id="target" slot="primary-action" part="button">
      <span slot="start" class="rb-icon"></span>
      <span class="rb-label"></span>
    </fluent-button>
    <fluent-menu-button slot="trigger" icon-only appearance="subtle" part="caret" aria-label="Show options" aria-haspopup="true" aria-expanded="false"></fluent-menu-button>
    <fluent-menu-list focusgroup="menu" popover part="list"></fluent-menu-list>
  </fluent-menu>
  ${TOOLTIP_PART}`;

/**
 * `<docen-ribbon-split-button icon="paste" label="Paste" event="paste" items='[{...}]'>`
 * — a split command (Office "Split Button"): a `fluent-button` primary action
 * (icon + label) plus an icon-only `fluent-menu-button` trigger (Fluent's
 * caret) that opens a native menu. The primary action emits `command { event }`;
 * a menu item emits `command { event, value }`.
 *
 * `size="large"` switches to the Office large split layout: the primary stacks
 * icon over label (column) and the caret drops below (fluent-menu data-vertical).
 */
class DocenRibbonSplitButton extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["label", "icon", "event", "items", "tooltip", "size", "icon-only"];
  }

  readonly #anchorId = `--rb-split-${++seq}`;
  #primary?: HTMLElement;
  #menu?: HTMLElement;
  #list?: HTMLElement;
  #icon?: HTMLSpanElement;
  #fwdCleanup?: () => void;
  #focusCleanup?: () => void;

  attributeChangedCallback(name: string): void {
    if (!this.shadowRoot) return;
    if (name === "icon") this.#renderIcon();
    if (name === "label" || name === "tooltip") this.#renderLabel();
    if (name === "icon-only") {
      this.#renderLabel();
      this.#syncIconSlot();
    }
    if (name === "items") this.#renderItems();
    if (name === "size") this.#applySize();
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    this.#primary = this.shadowRoot!.querySelector('fluent-button[slot="primary-action"]')!;
    this.#menu = this.shadowRoot!.querySelector("fluent-menu")!;
    this.#list = this.shadowRoot!.querySelector("fluent-menu-list")!;
    // Expose the primary action (the button's left edge) as a CSS anchor for
    // icon-only splits to left-align their dropdown to (see #applySize). The
    // primary lives in this shadow root, so the anchor resolves same-shadow
    // (like the combobox). Anchoring the host element itself crosses the shadow
    // boundary and the browser drops the popover at the viewport corner.
    this.#primary.style.anchorName = this.#anchorId;
    this.#icon = createIconSlot();
    this.#primary.prepend(this.#icon);
    this.#syncIconSlot();
    // Split: the primary action is a standalone invoke, not a menu toggle.
    this.#primary.addEventListener("click", (event) => {
      event.stopPropagation();
      this.#emitMain();
    });
    this.#renderIcon();
    this.#renderLabel();
    this.#renderItems();
    this.#applySize();
    // Split primary is flat by default (subtle); the split's hover border is
    // driven by fluent-menu (see registry). A caller can override via the
    // `appearance` attribute. The caret stays subtle.
    this.#fwdCleanup = forwardAttributes(this, this.#primary, { appearance: "subtle" });
    // Keep the editor's selection when the primary action is clicked — the caret
    // dropdown still opens normally (only the primary, not the whole host, is
    // guarded, so the menu trigger keeps its focus-driven open/keyboard behavior).
    this.#focusCleanup = preventFocusLoss(this.#primary);
  }

  disconnectedCallback(): void {
    this.#fwdCleanup?.();
    this.#focusCleanup?.();
  }

  get event(): string {
    return this.getAttribute("event") ?? this.getAttribute("label") ?? "";
  }

  get items(): RibbonMenuItem[] {
    try {
      return JSON.parse(this.getAttribute("items") ?? "[]") as RibbonMenuItem[];
    } catch {
      return [];
    }
  }

  #applySize(): void {
    if (!this.#menu) return;
    // data-vertical flips fluent-menu's split to a column (caret below) — see
    // the registry's fluent-menu compose. Pinned to size="large".
    if (this.getAttribute("size") === "large") this.#menu.setAttribute("data-vertical", "");
    else this.#menu.removeAttribute("data-vertical");
    // Fluent right-aligns a split's dropdown to the caret (inset-inline-end:
    // anchor(self-end)), so the dropdown fans out to the left. Whether that
    // reads as "under the button" or "pushed off to the left" depends on where
    // the button sits in the viewport — a right-edge button (e.g. Find) keeps
    // the default right-align and its dropdown swings far left, while a
    // left-edge button (e.g. Paste) ends up left-aligned once the browser pulls
    // the overflow back on-screen. Different splits therefore align
    // differently. Anchor the primary action (the button's left edge) and
    // left-align for every split — large and icon-only alike — so all
    // dropdowns line up under their button consistently (Office behavior).
    // Inline styles win over Fluent's ::slotted([popover]) positioning.
    if (this.#list) {
      this.#list.style.positionAnchor = this.#anchorId;
      this.#list.style.insetInlineStart = "anchor(self-start)";
      this.#list.style.insetInlineEnd = "unset";
    }
    // The tooltip (TOOLTIP_PART) anchors to #target (the primary) too, but its
    // connectedCallback runs before this host's (shadow children connect first),
    // so it reads #primary.anchorName while still empty and falls back to
    // "--target". #applySize then sets the primary's anchor-name to this
    // instance's --rb-split-N — leaving the tooltip's "--target" with no
    // matching anchor, so it strands at the viewport corner. Re-point the
    // tooltip at the same anchor name the primary now advertises.
    const tooltip = this.shadowRoot?.querySelector("fluent-tooltip") as HTMLElement | null;
    if (tooltip) tooltip.style.positionAnchor = this.#anchorId;
  }

  #tipText(): string {
    return this.getAttribute("tooltip") ?? this.getAttribute("label") ?? "";
  }

  /** Move the icon to the centered default slot when icon-only (matching
   *  <docen-ribbon-button>), else into Fluent's 'start' slot. */
  #syncIconSlot(): void {
    if (this.#icon) this.#icon.slot = this.hasAttribute("icon-only") ? "" : "start";
  }

  #renderIcon(): void {
    if (this.#icon) renderIcon(this.#icon, this.getAttribute("icon") ?? "");
  }

  #renderLabel(): void {
    // icon-only: hide the visible label (it still feeds the tooltip).
    const visible = this.hasAttribute("icon-only") ? "" : (this.getAttribute("label") ?? "");
    renderLabel(this.shadowRoot!.querySelector(".rb-label")!, visible);
    renderLabel(this.shadowRoot!.querySelector(".rb-tip")!, this.#tipText());
  }

  #renderItems(): void {
    const list = this.shadowRoot?.querySelector("fluent-menu-list");
    if (!list) return;
    list.replaceChildren();
    for (const item of this.items) {
      const menuItem = document.createElement("fluent-menu-item");
      menuItem.setAttribute("role", "menuitem");
      menuItem.setAttribute("data-indent", "0");
      menuItem.setAttribute("data-fg-ati", "0");
      menuItem.textContent = item.text;
      if (item.disabled) menuItem.setAttribute("disabled", "");
      menuItem.addEventListener("change", () => this.#emitItem(item));
      list.append(menuItem);
    }
  }

  #emitMain(): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.event, source: this },
      }),
    );
  }

  #emitItem(item: RibbonMenuItem): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: {
          // A split's menu item is a variant of the split's own command — the
          // command name is the split's event (this.event) unless the item sets
          // its own; the variant rides in value. Using item.value as the event
          // dropped the split context, so no host could route page-size /
          // margins / orientation / … .
          event: item.event ?? this.event,
          value: item.value,
          source: this,
        },
      }),
    );
  }
}

customElements.define("docen-ribbon-split-button", DocenRibbonSplitButton);

export default DocenRibbonSplitButton;
