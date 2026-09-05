import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string };

/** A `fluent-checkbox`. The rendered state follows `checked`; `currentChecked`
 *  is a separate slot only user clicks keep in sync. */
type FluentCheckbox = HTMLElement & { checked: boolean };

/** The dialog's values — the column count, the gap in centimeters, the
 *  separator-line flag, and the equal-width flag. All twip conversion stays
 *  on the host. */
export interface ColumnsValues {
  count: number;
  space: number;
  separate: boolean;
  equalWidth: boolean;
}

/** Word's defaults when a prefill field is absent (1 column, 0.5" gap). */
const DEFAULTS = { count: 1, space: 1.27 } as const;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .columns-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
    flex: 1 1 0;
    min-width: 0;
  }
  .field > label {
    white-space: nowrap;
  }
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
  .unit {
    white-space: nowrap;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 8px;
  }
`;

const template = html<DocenColumnsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="columns-body">
      <div class="row">
        <div class="field">
          <label ${ref("countLabel")}></label>
          <fluent-text-input ${ref("countInput")} type="number" min="1" max="9"></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("spaceLabel")}></label>
          <fluent-text-input
            ${ref("spaceInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
      <label class="check-field">
        <fluent-checkbox part="separate" ${ref("separateCheck")}></fluent-checkbox>
        <span ${ref("separateLabel")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox part="equal" ${ref("equalCheck")}></fluent-checkbox>
        <span ${ref("equalLabel")}></span>
      </label>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyColumns()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-columns-dialog>` — the Word "More Columns" dialog fields (column
 * count, spacing, separator line, equal widths). Opened by the Columns menu's
 * More Columns item; the host prefills from the current section via
 * `show(values)` and commits via `columns:ok`. Per-column manual widths stay
 * out until the host generates them. Rides on `<docen-dialog>` for the modal
 * shell.
 */
@customElement({ name: "docen-columns-dialog", template, styles })
class DocenColumnsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable countLabel?: HTMLElement;
  @observable countInput?: FluentTextInput;
  @observable spaceLabel?: HTMLElement;
  @observable spaceInput?: FluentTextInput;
  @observable separateCheck?: FluentCheckbox;
  @observable separateLabel?: HTMLElement;
  @observable equalCheck?: FluentCheckbox;
  @observable equalLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Prefill from the current section's w:cols (count/space only — the flags
   *  read straight off the checkboxes); absent values take Word's defaults. */
  show(values: Partial<ColumnsValues> = {}): void {
    const count = Number.isFinite(values.count) ? Math.max(1, values.count!) : DEFAULTS.count;
    if (this.countInput) this.countInput.value = String(count);
    if (this.spaceInput) {
      this.spaceInput.value =
        typeof values.space === "number" && Number.isFinite(values.space)
          ? String(Math.round(values.space * 100) / 100)
          : String(DEFAULTS.space);
    }
    if (this.separateCheck) this.separateCheck.checked = values.separate === true;
    if (this.equalCheck) this.equalCheck.checked = values.equalWidth !== false;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyColumns(): void {
    const n = Number(this.countInput?.value);
    // Word caps the count per paper width; the host re-clamps to 9 either way.
    const count = Number.isFinite(n) && n >= 1 ? Math.min(9, Math.trunc(n)) : DEFAULTS.count;
    const s = Number(this.spaceInput?.value);
    // A cleared or garbage gap keeps the default rather than committing 0.
    const space = Number.isFinite(s) && s > 0 ? s : DEFAULTS.space;
    this.$emit("columns:ok", {
      count,
      space,
      separate: this.separateCheck?.checked === true,
      equalWidth: this.equalCheck?.checked !== false,
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("columnsDialog.title", this);
    if (this.countLabel) this.countLabel.textContent = t("columnsDialog.count", this);
    if (this.spaceLabel) this.spaceLabel.textContent = t("columnsDialog.space", this);
    if (this.separateLabel) this.separateLabel.textContent = t("columnsDialog.separate", this);
    if (this.equalLabel) this.equalLabel.textContent = t("columnsDialog.equalWidth", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".unit") ?? [])
      el.textContent = t("pageSetup.cm", this);
  }
}

export default DocenColumnsDialog;
