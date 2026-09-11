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
  /* size="large" — Office large button: icon stacked over label (column),
     larger glyph. Same geometry as the plain button's large size (70px row,
     104px label cap) so the two sit level in one group row. */
  :host([size="large"]) {
    flex-shrink: 0;
  }
  :host([size="large"]) fluent-toggle-button {
    flex-direction: column;
    justify-content: flex-start;
    min-width: 0;
    max-width: 104px;
    min-height: 70px;
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
    /* A mixed Latin+CJK label ("Markdown 输入") wraps at the space, never
       mid-CJK-word. */
    word-break: keep-all;
    overflow-wrap: break-word;
  }
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

  // pressed is synced imperatively (not via binding): the host flag is the one
  // truth, and a one-way binding would race the sync on every transaction.
  pressedChanged(): void {
    this.syncPressed();
  }

  // Fluent's own press() self-flip is a second writer fighting the sync — it
  // fires after the command round-trip and re-flips the value the host just
  // stamped, leaving the lit state out of phase with the flag. Neutralize it
  // per instance so syncPressed stays the only writer of Fluent's pressed.
  toggleBtnChanged(): void {
    if (this.toggleBtn) (this.toggleBtn as { press?: () => void }).press = () => {};
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
    // Defer past the rest of this click's listeners so the host's sync write
    // doesn't re-enter the same event dispatch.
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
