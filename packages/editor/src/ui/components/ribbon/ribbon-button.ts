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
  /* Labelled button (default = small): Fluent's button font is 14px
     (fontSizeBase300) and min-height 32px — too tall for a compact ribbon
     row. Drop the label to 12px and the button to 26px so a 3-row rb-grid of
     small commands stays tight. size="large" overrides both below; icon-only
     hides the label. size="small" is the same as the default, for callers
     that want an explicit, uniform size attribute across a group. */
  fluent-button {
    min-height: 26px;
  }
  .rb-label,
  :host([size="small"]) .rb-label {
    font-size: 12px;
  }
  /* size="large" — Office large button: icon stacked over label (column),
     larger glyph. Its height matches a large split (primary 56 + caret 14 =
     70) and its content is top-aligned (justify-content:flex-start), so the
     icon/label line up with a sibling large split's primary — the bottom gap
     stands in for the split's caret row. */
  :host([size="large"]) {
    flex-shrink: 0;
  }
  :host([size="large"]) fluent-button {
    flex-direction: column;
    justify-content: flex-start;
    min-width: 0;
    max-width: 90px;
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
    overflow-wrap: break-word;
  }
`;

const template = html<DocenRibbonButton>`
  <fluent-button
    id="target"
    part="button"
    appearance="${(x) => x.appearance ?? "subtle"}"
    ?disabled="${(x) => x.disabled}"
    ?icon-only="${(x) => x.iconOnly}"
    @click="${(x) => x.onClick()}"
  >
    <span class="rb-icon" ${ref("iconSlot")}></span>
    <span class="rb-label">${(x) => x.visibleLabel}</span>
  </fluent-button>
  <fluent-tooltip anchor="target" positioning="top">
    <span class="rb-tip">${(x) => x.tooltipText}</span>
  </fluent-tooltip>
`;

/**
 * `<docen-ribbon-button icon="cut" label="Cut" event="cut">` — a command that
 * wraps `<fluent-button appearance="subtle">` with a built-in `<fluent-tooltip>`
 * (label, or a `tooltip` attr override). The icon (a known Office name) goes
 * into `slot="start"`; with `icon-only` the label is hidden but still surfaces
 * as the tooltip. Click emits `command` with `{ event }`.
 */
@customElement({ name: "docen-ribbon-button", template, styles })
class DocenRibbonButton extends FASTElement {
  // Optional (no initializer): under useDefineForClassFields an initializer
  // would shadow the @attr-installed getter/setter and break reactivity.
  @attr label?: string;
  @attr icon?: string;
  @attr event?: string;
  @attr tooltip?: string;
  @attr appearance?: string;
  @attr({ mode: "boolean" }) disabled?: boolean;
  @attr({ attribute: "icon-only", mode: "boolean" }) iconOnly?: boolean;

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

  iconChanged(): void {
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
  }

  iconOnlyChanged(): void {
    this.syncIconSlot();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.syncIconSlot();
    if (this.iconSlot) renderIcon(this.iconSlot, this.icon ?? "");
    // Keep the editor's selection on click — mousedown preventDefault (capture)
    // stops the contenteditable from blurring before the command runs, mirroring
    // Tiptap BubbleMenu's handler.
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

  onClick(): void {
    if (this.disabled) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.eventName, source: this },
      }),
    );
  }
}

export default DocenRibbonButton;
