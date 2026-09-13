import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string };

/** A `fluent-dropdown` combobox: `value` reads/writes the picked option. */
type FluentDropdown = HTMLElement & { value: string };

/** The Line Numbering Options dialog's values — distance in centimeters
 *  (undefined = Word's 自动 margin placement); the twip conversion stays on
 *  the host. */
export interface LineNumbersValues {
  /** First line number after each restart (w:lnNumType start). */
  start: number;
  /** Number only every Nth line (w:lnNumType countBy). */
  countBy: number;
  /** Gap from the text margin in cm (w:lnNumType distance); undefined =
   *  Word's auto placement. */
  distance?: number;
  /** When the counter resets (w:lnNumType restart). */
  restart: "continuous" | "newPage" | "newSection";
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .body {
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
`;

const template = html<DocenLineNumbersDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <div class="field">
          <label ${ref("startLabel")}></label>
          <fluent-text-input
            ${ref("startInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("distanceLabel")}></label>
          <fluent-text-input
            ${ref("distanceInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("countByLabel")}></label>
          <fluent-text-input
            ${ref("countByInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("restartLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("restartDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="continuous"></fluent-option>
              <fluent-option value="newPage"></fluent-option>
              <fluent-option value="newSection"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyOptions()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-line-numbers-dialog>` — Word's "Line Numbering Options" fields
 * (the Layout tab's Line Numbers menu options entry): start number, distance
 * from text, count stride, and the restart mode. Opened via `show(values)`
 * from the current section's w:lnNumType and committed via `line-numbers:ok`.
 */
@customElement({ name: "docen-line-numbers-dialog", template, styles })
class DocenLineNumbersDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable startLabel?: HTMLElement;
  @observable distanceLabel?: HTMLElement;
  @observable countByLabel?: HTMLElement;
  @observable restartLabel?: HTMLElement;
  @observable startInput?: FluentTextInput;
  @observable distanceInput?: FluentTextInput;
  @observable countByInput?: FluentTextInput;
  @observable restartDropdown?: FluentDropdown;
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

  /** Prefill from the current section's w:lnNumType; absent fields fall back
   *  to Word's defaults (start 1, count 1, auto distance, continuous). */
  show(values: Partial<LineNumbersValues> = {}): void {
    if (this.startInput) this.startInput.value = String(values.start ?? 1);
    if (this.distanceInput)
      this.distanceInput.value = typeof values.distance === "number" ? String(values.distance) : "";
    if (this.countByInput) this.countByInput.value = String(values.countBy ?? 1);
    if (this.restartDropdown) this.restartDropdown.value = values.restart ?? "continuous";
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). A cleared
   *  distance field is Word's 自动 — undefined. */
  applyOptions(): void {
    const distance = Number(this.distanceInput?.value);
    this.$emit("line-numbers:ok", {
      start: Math.max(1, Math.round(Number(this.startInput?.value) || 1)),
      countBy: Math.max(1, Math.round(Number(this.countByInput?.value) || 1)),
      distance:
        Number.isFinite(distance) && distance > 0 ? Math.round(distance * 100) / 100 : undefined,
      // The dropdown only offers its own tokens — the value reads back
      // closed, no wider validation needed.
      restart: (this.restartDropdown?.value ?? "continuous") as LineNumbersValues["restart"],
    } satisfies LineNumbersValues);
    this.hide();
  }

  #labelOptions(dropdown: FluentDropdown | undefined, labels: string[]): void {
    if (!dropdown) return;
    dropdown.querySelectorAll("fluent-option").forEach((opt, i) => {
      if (labels[i]) opt.textContent = labels[i];
    });
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("lineNumbers.title", this);
    if (this.startLabel) this.startLabel.textContent = t("lineNumbers.start", this);
    if (this.distanceLabel) this.distanceLabel.textContent = t("lineNumbers.distance", this);
    if (this.distanceInput)
      this.distanceInput.setAttribute("placeholder", t("lineNumbers.auto", this));
    if (this.countByLabel) this.countByLabel.textContent = t("lineNumbers.countBy", this);
    if (this.restartLabel) this.restartLabel.textContent = t("lineNumbers.restart", this);
    // The menu's own entries name the same three modes — reuse their words.
    this.#labelOptions(this.restartDropdown, [
      t("ribbon.opt.continuous-line-numbers", this),
      t("ribbon.opt.restart-each-page", this),
      t("ribbon.opt.restart-each-section", this),
    ]);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".unit") ?? [])
      el.textContent = t("pageSetup.cm", this);
  }
}

export default DocenLineNumbersDialog;
