import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../ui/i18n/localize";

/** A selectable option for a `radio` field. */
export interface PropertyOption {
  readonly label: string;
  readonly value: string;
}

/** A single editable property rendered inside a group. */
export interface PropertyField {
  /** Control kind: `radio` renders an option list; `number`/`color` a label + value. */
  readonly type: "radio" | "number" | "color";
  /** Stable id emitted on `property:change`. */
  readonly key: string;
  /** Row label (number/color) or group heading (radio, optional). */
  readonly label?: string;
  /** Current value (the selected option's value for radio). */
  readonly value?: string;
  /** radio options. Required for `type: "radio"`. */
  readonly options?: readonly PropertyOption[];
}

/** A collapsible group of property fields. */
export interface PropertyGroup {
  readonly title: string;
  readonly fields: readonly PropertyField[];
  /** Expanded by default. */
  readonly expanded?: boolean;
}

const styles = css`
  :host {
    display: block;
    font-size: 12px;
  }
  fluent-accordion,
  fluent-accordion-item {
    max-width: 100%;
    width: 100%;
  }
  fluent-accordion-item::part(heading) {
    font-size: 12px;
    font-weight: 600;
  }
  fluent-accordion-item::part(content) {
    padding-inline-start: 28px;
    padding-inline-end: 6px;
  }
  /* A row is "label on the inline-start, control on the inline-end" —
     space-between pins the label left and the control right without the
     control stretching. */
  .field-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    padding: 4px 0;
  }
  .field-row label,
  .field-label {
    font-size: 12px;
    color: #666;
  }
  .field-label {
    display: block;
    padding: 4px 0;
  }
  .field-row fluent-text-input {
    flex: 0 0 auto;
    width: 88px;
  }
  .color-swatch {
    flex: 0 0 auto;
    width: 56px;
    height: 24px;
    border: 1px solid #d0d0d0;
    border-radius: 4px;
    padding: 2px;
    background: #fff;
    cursor: pointer;
  }
  .empty {
    font-size: 12px;
    color: #666;
    padding: 12px 8px;
    margin: 0;
  }
`;

const template = html<DocenFormatPane>`<div part="body" ${ref("body")}></div>`;

/**
 * `<docen-format-pane groups='[{title, fields:[…]}]'>` — an Office-style format
 * pane: a Fluent accordion where each group holds typed fields (`radio` /
 * `number` / `color`). Editing a field emits `property:change` with
 * `{ key, value }`. Field labels and option text are business strings
 * (translated by the editor package); only the empty state is localized here.
 */
@customElement({ name: "docen-format-pane", template, styles })
class DocenFormatPane extends FASTElement {
  @attr groups?: string;
  @observable body?: HTMLElement;
  #unsubscribe?: () => void;

  groupsChanged(): void {
    this.#render();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#render();
    // Only the empty-state string is component-internal; field/option labels
    // are business strings, so update just the empty line on lang change.
    this.#unsubscribe = observeLang(() => {
      const empty = this.body?.querySelector(".empty");
      if (empty) empty.textContent = t("properties.empty", this);
    });
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  // The JSON `groups` attribute parsed; bad JSON → empty (no render).
  get parsedGroups(): PropertyGroup[] {
    try {
      return JSON.parse(this.groups ?? "[]") as PropertyGroup[];
    } catch {
      return [];
    }
  }

  #render(): void {
    const body = this.body;
    if (!body) return;
    body.replaceChildren();
    const groups = this.parsedGroups;
    if (groups.length === 0) {
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = t("properties.empty", this);
      body.append(empty);
      return;
    }
    const accordion = document.createElement("fluent-accordion");
    accordion.setAttribute("expand-mode", "multi");
    for (const group of groups) {
      accordion.append(this.#renderGroup(group));
    }
    body.append(accordion);
  }

  #renderGroup(group: PropertyGroup): HTMLElement {
    const item = document.createElement("fluent-accordion-item");
    if (group.expanded !== false) item.setAttribute("expanded", "");
    const heading = document.createElement("span");
    heading.slot = "heading";
    heading.textContent = group.title;
    item.append(heading);
    for (const field of group.fields) {
      item.append(this.#renderField(field));
    }
    return item;
  }

  #renderField(field: PropertyField): Node {
    if (field.type === "radio") return this.#renderRadio(field);
    const row = document.createElement("div");
    row.className = "field-row";
    const label = document.createElement("label");
    label.textContent = field.label ?? "";
    row.append(label);
    if (field.type === "number") {
      const input = document.createElement("fluent-text-input") as HTMLElement & {
        value: string;
      };
      input.setAttribute("type", "number");
      if (field.value != null) input.setAttribute("value", field.value);
      input.addEventListener("input", () => this.#emit(field.key, input.value));
      row.append(input);
    } else {
      const input = document.createElement("input");
      input.type = "color";
      input.className = "color-swatch";
      if (field.value != null) input.value = field.value;
      input.addEventListener("input", () => this.#emit(field.key, input.value));
      row.append(input);
    }
    return row;
  }

  #renderRadio(field: PropertyField): HTMLElement {
    const wrap = document.createElement("div");
    if (field.label) {
      const heading = document.createElement("span");
      heading.className = "field-label";
      heading.textContent = field.label;
      wrap.append(heading);
    }
    const group = document.createElement("fluent-radio-group");
    group.setAttribute("name", field.key);
    group.setAttribute("orientation", "vertical");
    for (const option of field.options ?? []) {
      const id = `${field.key}-${option.value}`;
      const fieldEl = document.createElement("fluent-field");
      fieldEl.setAttribute("label-position", "after");
      const label = document.createElement("label");
      label.slot = "label";
      label.htmlFor = id;
      label.id = `${id}--label`;
      label.textContent = option.label;
      const radio = document.createElement("fluent-radio") as HTMLElement & {
        value: string;
      };
      radio.slot = "input";
      radio.id = id;
      radio.setAttribute("name", field.key);
      radio.value = option.value;
      radio.setAttribute("aria-labelledby", `${id}--label`);
      if (field.value === option.value) radio.setAttribute("checked", "");
      radio.addEventListener("change", () => this.#emit(field.key, option.value));
      fieldEl.append(label, radio);
      group.append(fieldEl);
    }
    wrap.append(group);
    return wrap;
  }

  #emit(key: string, value: string): void {
    this.dispatchEvent(
      new CustomEvent("property:change", {
        bubbles: true,
        composed: true,
        detail: { key, value, source: this },
      }),
    );
  }
}

export default DocenFormatPane;
