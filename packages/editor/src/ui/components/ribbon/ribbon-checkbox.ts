import { FASTElement, attr, css, customElement, html } from "@microsoft/fast-element";

import { COMMAND_HOST_STYLE } from "./command-helpers";

const styles = css`
  ${COMMAND_HOST_STYLE}
  /* Word's Table Style Options row: a compact checkbox line. The label span
     sits outside fluent-checkbox — its template has no default label slot
     (only the indicator slots), the same wrapping the find-replace dialog
     uses; the wrapping <label> keeps clicks routing to the box. */
  label.rb-check {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
    min-height: 20px;
  }
  .rb-label {
    font-size: 12px;
  }
`;

const template = html<DocenRibbonCheckbox>`
  <label class="rb-check">
    <fluent-checkbox
      part="checkbox"
      ?checked="${(x) => x.checked}"
      ?disabled="${(x) => x.disabled}"
      @change="${(x) => x.onChange()}"
    ></fluent-checkbox>
    <span class="rb-label">${(x) => x.visibleLabel}</span>
  </label>
`;

/**
 * `<docen-ribbon-checkbox label="Header Row" event="toggle-table-look">` — a
 * labelled checkbox command (Office.js manifest `CheckBox`): Word's Table
 * Style Options flags. Toggling emits `command` with `{ event, value }`; the
 * `checked` attr is the visual stamp (the host re-stamps it to mirror the
 * live table state).
 */
@customElement({ name: "docen-ribbon-checkbox", template, styles })
class DocenRibbonCheckbox extends FASTElement {
  // Optional (no initializer): under useDefineForClassFields an initializer
  // would shadow the @attr-installed getter/setter and break reactivity.
  @attr label?: string;
  @attr event?: string;
  @attr value?: string;
  @attr({ mode: "boolean" }) checked?: boolean;
  @attr({ mode: "boolean" }) disabled?: boolean;

  get visibleLabel(): string {
    return this.label ?? "";
  }
  get eventName(): string {
    return this.event || this.label || "";
  }

  onChange(): void {
    if (this.disabled) return;
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: this.eventName, value: this.value, source: this },
      }),
    );
  }
}

export default DocenRibbonCheckbox;
