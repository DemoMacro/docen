import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import {
  COMMAND_HOST_STYLE,
  appendMenuItems,
  renderIcon,
  suppressTooltipWhileMenuOpen,
} from "./command-helpers";
import type { RibbonMenuItem } from "./ribbon-menu";

// Per-instance CSS anchor name so each split's dropdown aligns to its own
// button (left edge), not Fluent's default caret right-edge — which on a
// narrow icon-only split pushes the dropdown far to the left.
let seq = 0;

const styles = css`
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
  :host([size="large"]) {
    flex-shrink: 0;
  }
  :host([size="large"]) fluent-button {
    flex-direction: column;
    justify-content: flex-start;
    min-width: 0;
    max-width: 90px;
    min-height: 56px;
    padding: 4px 12px;
  }
  :host([size="large"]) .rb-icon svg {
    width: 32px;
    height: 32px;
  }
  :host([size="large"]) .rb-label {
    font-size: 11px;
    text-align: center;
    line-height: 1.2;
    white-space: normal;
    overflow-wrap: break-word;
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
`;

const template = html<DocenRibbonSplitButton>`
  <fluent-menu split part="menu" style="--menu-max-height: auto;" ${ref("menu")}>
    <fluent-button
      id="target"
      slot="primary-action"
      part="button"
      appearance="${(x) => x.appearance ?? "subtle"}"
      ?disabled="${(x) => x.disabled}"
      ${ref("primary")}
    >
      <span slot="start" class="rb-icon" ${ref("iconSlot")}></span>
      <span class="rb-label">${(x) => x.visibleLabel}</span>
    </fluent-button>
    <fluent-menu-button
      slot="trigger"
      icon-only
      appearance="subtle"
      part="caret"
      aria-label="Show options"
      aria-haspopup="true"
      aria-expanded="false"
      ?disabled="${(x) => x.disabled}"
    ></fluent-menu-button>
    <fluent-menu-list focusgroup="menu" popover part="list" ${ref("list")}></fluent-menu-list>
  </fluent-menu>
  <fluent-tooltip anchor="target" positioning="top" ${ref("tooltipEl")}>
    <span class="rb-tip">${(x) => x.tooltipText}</span>
  </fluent-tooltip>
`;

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
@customElement({ name: "docen-ribbon-split-button", template, styles })
class DocenRibbonSplitButton extends FASTElement {
  @attr label?: string;
  @attr icon?: string;
  @attr event?: string;
  @attr items?: string;
  @attr tooltip?: string;
  @attr size?: string;
  @attr appearance?: string;
  @attr({ mode: "boolean" }) disabled?: boolean;
  @attr({ attribute: "icon-only", mode: "boolean" }) iconOnly?: boolean;

  @observable primary?: HTMLElement;
  @observable menu?: HTMLElement;
  @observable list?: HTMLElement;
  @observable iconSlot?: HTMLSpanElement;
  @observable tooltipEl?: HTMLElement;

  readonly anchorId = `--rb-split-${++seq}`;

  #tooltipDisposer?: () => void;

  get eventName(): string {
    return this.event || this.label || "";
  }
  get visibleLabel(): string {
    return this.iconOnly ? "" : (this.label ?? "");
  }
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
  iconOnlyChanged(): void {
    this.syncIconSlot();
  }
  itemsChanged(): void {
    this.renderItems();
  }
  sizeChanged(): void {
    this.applySize();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.syncIconSlot();
    this.applyAnchor();
    this.applySize();
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
    this.renderItems();
    // Split: the primary action is a standalone invoke, not a menu toggle —
    // stopPropagation keeps the click from opening the caret dropdown.
    this.primary?.addEventListener("click", this.onPrimaryClick);
    // Keep the editor's selection when the primary action is clicked — the
    // caret dropdown still opens normally (only the primary is guarded).
    this.primary?.addEventListener("mousedown", this.onMousedown, { capture: true });
    this.#tooltipDisposer = suppressTooltipWhileMenuOpen(this.tooltipEl, this.list);
  }

  disconnectedCallback(): void {
    this.primary?.removeEventListener("click", this.onPrimaryClick);
    this.primary?.removeEventListener("mousedown", this.onMousedown, { capture: true });
    this.#tooltipDisposer?.();
    this.#tooltipDisposer = undefined;
    super.disconnectedCallback();
  }

  private readonly onMousedown = (event: Event): void => event.preventDefault();

  /** Move the icon to the centered default slot when icon-only (matching
   *  <docen-ribbon-button>), else into Fluent's 'start' slot. */
  private syncIconSlot(): void {
    if (this.iconSlot) this.iconSlot.slot = this.iconOnly ? "" : "start";
  }

  private applyAnchor(): void {
    // Expose the primary action (left edge) as a CSS anchor for icon-only
    // splits to left-align their dropdown to. The primary lives in this shadow
    // root, so the anchor resolves same-shadow. Anchoring the host itself
    // crosses the shadow boundary and the browser drops the popover at the
    // viewport corner.
    if (this.primary) this.primary.style.anchorName = this.anchorId;
    if (this.list) {
      this.list.style.positionAnchor = this.anchorId;
      this.list.style.insetInlineStart = "anchor(self-start)";
      this.list.style.insetInlineEnd = "unset";
    }
    // The tooltip anchors to #target (the primary) but connects before this
    // host, so re-point it at the same anchor the primary now advertises.
    if (this.tooltipEl) this.tooltipEl.style.positionAnchor = this.anchorId;
  }

  private applySize(): void {
    if (!this.menu) return;
    // data-vertical flips fluent-menu's split to a column (caret below) — see
    // the registry's fluent-menu compose. Pinned to size="large".
    if (this.size === "large") this.menu.setAttribute("data-vertical", "");
    else this.menu.removeAttribute("data-vertical");
  }

  private renderItems(): void {
    if (this.list) appendMenuItems(this.list, this.parsedItems, (item) => this.emitItem(item));
  }

  private readonly onPrimaryClick = (event: Event): void => {
    event.stopPropagation();
    if (this.disabled) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.eventName, source: this },
      }),
    );
  };

  private emitItem(item: RibbonMenuItem): void {
    if (this.disabled) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: {
          // A split's menu item is a variant of the split's own command — the
          // command name is the split's event unless the item sets its own;
          // the variant rides in value.
          event: item.event ?? this.eventName,
          value: item.value,
          source: this,
        },
      }),
    );
  }
}

export default DocenRibbonSplitButton;
