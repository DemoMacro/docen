import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** The dialog's values in centimeters — the unit Word's zh dialogs show. All
 *  conversion to/from OOXML twips stays on the host. */
export interface PageSetupValues {
  margins: { top: number; bottom: number; left: number; right: number };
  size: { width: number; height: number };
}

/** Word's defaults (Normal margins on A4) for absent prefill fields. */
const DEFAULTS = {
  margin: 2.54,
  side: 3.18,
  width: 21,
  height: 29.7,
} as const;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(400px, 92vw);
  }
  .setup-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .setup-heading {
    font-weight: 600;
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

const template = html<DocenPageSetupDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="setup-body">
      <div class="setup-heading" ${ref("marginsHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("topLabel")}></label>
          <fluent-text-input
            ${ref("topInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
        <div class="field">
          <label ${ref("bottomLabel")}></label>
          <fluent-text-input
            ${ref("bottomInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("leftLabel")}></label>
          <fluent-text-input
            ${ref("leftInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
        <div class="field">
          <label ${ref("rightLabel")}></label>
          <fluent-text-input
            ${ref("rightInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
      <div class="setup-heading" ${ref("sizeHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("widthLabel")}></label>
          <fluent-text-input
            ${ref("widthInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
        <div class="field">
          <label ${ref("heightLabel")}></label>
          <fluent-text-input
            ${ref("heightInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applySetup()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-page-setup-dialog>` — the Word "Page Setup" geometry fields (margins
 * plus paper size, in centimeters). Opened by the Margins menu's Custom Margins
 * and the Size menu's More Paper Sizes items; the host prefills from the
 * current section via `show(values)` and commits via `page-setup:ok`. Gutter,
 * header/footer distance and the layout tabs stay out until the engine
 * consumes them. Rides on `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-page-setup-dialog", template, styles })
class DocenPageSetupDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable marginsHeading?: HTMLElement;
  @observable topLabel?: HTMLElement;
  @observable topInput?: FluentTextInput;
  @observable bottomLabel?: HTMLElement;
  @observable bottomInput?: FluentTextInput;
  @observable leftLabel?: HTMLElement;
  @observable leftInput?: FluentTextInput;
  @observable rightLabel?: HTMLElement;
  @observable rightInput?: FluentTextInput;
  @observable sizeHeading?: HTMLElement;
  @observable widthLabel?: HTMLElement;
  @observable widthInput?: FluentTextInput;
  @observable heightLabel?: HTMLElement;
  @observable heightInput?: FluentTextInput;
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

  /** Prefill every field from the current section's geometry in centimeters;
   *  absent values fall back to Word's defaults (Normal margins on A4). */
  show(
    values: {
      margins?: Partial<PageSetupValues["margins"]>;
      size?: Partial<PageSetupValues["size"]>;
    } = {},
  ): void {
    const margins = values.margins ?? {};
    const size = values.size ?? {};
    if (this.topInput) this.topInput.value = this.#cm(margins.top, DEFAULTS.margin);
    if (this.bottomInput) this.bottomInput.value = this.#cm(margins.bottom, DEFAULTS.margin);
    if (this.leftInput) this.leftInput.value = this.#cm(margins.left, DEFAULTS.side);
    if (this.rightInput) this.rightInput.value = this.#cm(margins.right, DEFAULTS.side);
    if (this.widthInput) this.widthInput.value = this.#cm(size.width, DEFAULTS.width);
    if (this.heightInput) this.heightInput.value = this.#cm(size.height, DEFAULTS.height);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applySetup(): void {
    const margins = {
      top: this.#num(this.topInput?.value, DEFAULTS.margin),
      bottom: this.#num(this.bottomInput?.value, DEFAULTS.margin),
      left: this.#num(this.leftInput?.value, DEFAULTS.side),
      right: this.#num(this.rightInput?.value, DEFAULTS.side),
    };
    const size = {
      width: this.#num(this.widthInput?.value, DEFAULTS.width),
      height: this.#num(this.heightInput?.value, DEFAULTS.height),
    };
    this.$emit("page-setup:ok", { margins, size });
    this.hide();
  }

  /** Twips-sourced centimeters → the input's 2-decimal text. */
  #cm(value: number | undefined, fallback: number): string {
    if (typeof value !== "number" || !Number.isFinite(value)) return String(fallback);
    return String(Math.round(value * 100) / 100);
  }

  #num(v: string | undefined, fallback: number): number {
    const n = Number(v);
    // A cleared or garbage field keeps the prefill rather than committing 0.
    return Number.isFinite(n) && n > 0 ? n : fallback;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("ribbon.group.page-setup", this);
    if (this.marginsHeading) this.marginsHeading.textContent = t("ribbon.cmd.margins", this);
    if (this.sizeHeading) this.sizeHeading.textContent = t("ribbon.cmd.page-size", this);
    if (this.topLabel) this.topLabel.textContent = t("pageSetup.top", this);
    if (this.bottomLabel) this.bottomLabel.textContent = t("pageSetup.bottom", this);
    if (this.leftLabel) this.leftLabel.textContent = t("pageSetup.left", this);
    if (this.rightLabel) this.rightLabel.textContent = t("pageSetup.right", this);
    if (this.widthLabel) this.widthLabel.textContent = t("pageSetup.width", this);
    if (this.heightLabel) this.heightLabel.textContent = t("pageSetup.height", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    // The unit chips after each input share one text.
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".unit") ?? [])
      el.textContent = t("pageSetup.cm", this);
  }
}

export default DocenPageSetupDialog;
