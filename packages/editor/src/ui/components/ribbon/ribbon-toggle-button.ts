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

const styles = css`
  ${COMMAND_HOST_STYLE}
`;

const template = html<DocenRibbonToggleButton>`
  <fluent-toggle-button
    id="target"
    part="button"
    appearance="${(x) => x.appearance ?? "subtle"}"
    ?disabled="${(x) => x.disabled}"
    ?icon-only="${(x) => x.iconOnly}"
    @click="${(x) => x.onClick()}"
    ${ref("toggleBtn")}
  >
    <span class="rb-icon" ${ref("iconSlot")}></span>
    <span class="rb-label">${(x) => x.visibleLabel}</span>
  </fluent-toggle-button>
  <fluent-tooltip anchor="target" positioning="top">
    <span class="rb-tip">${(x) => x.tooltipText}</span>
  </fluent-tooltip>
`;

/**
 * `<docen-ribbon-toggle-button icon="bold" label="Bold" event="bold" icon-only>`
 * — a two-state command (Office "Toggle Button") wrapping
 * `<fluent-toggle-button appearance="subtle">` with a built-in tooltip. Fluent
 * owns the pressed affordance; `pressed` (attr) seeds the initial state and the
 * getter reads Fluent's live value. Click emits `command` with
 * `{ event, pressed }`.
 */
@customElement({ name: "docen-ribbon-toggle-button", template, styles })
class DocenRibbonToggleButton extends FASTElement {
  @attr label?: string;
  @attr icon?: string;
  @attr event?: string;
  @attr tooltip?: string;
  @attr appearance?: string;
  @attr({ mode: "boolean" }) pressed?: boolean;
  @attr({ mode: "boolean" }) disabled?: boolean;
  @attr({ attribute: "icon-only", mode: "boolean" }) iconOnly?: boolean;

  @observable toggleBtn?: HTMLElement;
  @observable iconSlot?: HTMLSpanElement;

  /** Icon-only hides the visible label (it still feeds the tooltip). */
  get visibleLabel(): string {
    return this.iconOnly ? "" : (this.label ?? "");
  }
  get tooltipText(): string {
    return this.tooltip || this.label || "";
  }
  get eventName(): string {
    return this.event || this.label || "";
  }
  /** Read Fluent's live pressed — the user toggles it directly on click. */
  get pressedState(): boolean {
    return (this.toggleBtn as { pressed?: boolean } | undefined)?.pressed ?? false;
  }

  iconChanged(): void {
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
  }

  iconOnlyChanged(): void {
    this.syncIconSlot();
  }

  // pressed is synced imperatively (not via binding): the user clicks Fluent's
  // button to flip its internal pressed, and a one-way binding would re-push
  // the stale host value on the next render and clobber that flip.
  pressedChanged(): void {
    this.syncPressed();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.syncIconSlot();
    this.syncPressed();
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
    // Keep the editor's selection on click — see ribbon-button.
    this.addEventListener("mousedown", this.onMousedown, { capture: true });
  }

  disconnectedCallback(): void {
    this.removeEventListener("mousedown", this.onMousedown, { capture: true });
    super.disconnectedCallback();
  }

  private readonly onMousedown = (event: Event): void => event.preventDefault();

  /** Center the glyph (default slot) when icon-only, else Fluent's 'start' slot. */
  private syncIconSlot(): void {
    if (this.iconSlot) this.iconSlot.slot = this.iconOnly ? "" : "start";
  }

  private syncPressed(): void {
    if (!this.toggleBtn) return;
    if (this.pressed) this.toggleBtn.setAttribute("pressed", "");
    else this.toggleBtn.removeAttribute("pressed");
  }

  onClick(): void {
    if (this.disabled) return;
    // Defer until Fluent has toggled its internal pressed state for this click.
    queueMicrotask(() => this.emit());
  }

  private emit(): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.eventName, pressed: this.pressedState, source: this },
      }),
    );
  }
}

export default DocenRibbonToggleButton;
