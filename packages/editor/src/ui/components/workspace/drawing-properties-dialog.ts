import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { DrawingPropertiesPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";

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
  .heading {
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

const template = html<DocenDrawingPropertiesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="heading" ${ref("sizeHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("widthLabel")}></label>
          <fluent-text-input ${ref("widthInput")} type="number" step="any"></fluent-text-input>
          <span class="unit">cm</span>
        </div>
        <div class="field">
          <label ${ref("heightLabel")}></label>
          <fluent-text-input ${ref("heightInput")} type="number" step="any"></fluent-text-input>
          <span class="unit">cm</span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("rotationLabel")}></label>
          <fluent-text-input ${ref("rotationInput")} type="number" step="any"></fluent-text-input>
          <span class="unit">°</span>
        </div>
      </div>
      <div class="heading" ${ref("positionHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("horizontalLabel")}></label>
          <fluent-text-input ${ref("horizontalInput")} type="number" step="any"></fluent-text-input>
          <span class="unit">cm</span>
        </div>
        <div class="field">
          <label ${ref("verticalLabel")}></label>
          <fluent-text-input ${ref("verticalInput")} type="number" step="any"></fluent-text-input>
          <span class="unit">cm</span>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyProperties()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** The prefill shape the host derives from the selected drawing: everything
 *  already in centimeters (the dialog's display unit), rotation in degrees. */
export interface DrawingPropertiesState {
  widthCm: number;
  heightCm: number;
  rotationDeg: number;
  offsetHCm: number;
  offsetVCm: number;
}

/**
 * `<docen-drawing-properties-dialog>` — the Word "Layout" dialog's numeric
 * core (Size and Position): the selected floating drawing's width/height,
 * rotation, and anchor offsets. The host prefills via `show(state)`; OK emits
 * `drawing-properties:ok` with a {@link DrawingPropertiesPatch} for the host
 * to stamp via `drawing-properties-apply`. Rides on `<docen-dialog>` for the
 * modal shell.
 */
@customElement({ name: "docen-drawing-properties-dialog", template, styles })
class DocenDrawingPropertiesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable sizeHeading?: HTMLElement;
  @observable widthLabel?: HTMLElement;
  @observable widthInput?: FluentTextInput;
  @observable heightLabel?: HTMLElement;
  @observable heightInput?: FluentTextInput;
  @observable rotationLabel?: HTMLElement;
  @observable rotationInput?: FluentTextInput;
  @observable positionHeading?: HTMLElement;
  @observable horizontalLabel?: HTMLElement;
  @observable horizontalInput?: FluentTextInput;
  @observable verticalLabel?: HTMLElement;
  @observable verticalInput?: FluentTextInput;
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

  /** Prefill from the selected drawing (the host's centimeter normalization:
   *  an image rides px attrs, a shape payload rides EMU — both land here as
   *  cm). Two decimals mirror Word's spin precision. */
  show(state: DrawingPropertiesState): void {
    const set = (input: FluentTextInput | undefined, cm: number): void => {
      if (input) input.value = String(Math.round(cm * 100) / 100);
    };
    set(this.widthInput, state.widthCm);
    set(this.heightInput, state.heightCm);
    set(this.rotationInput, state.rotationDeg);
    set(this.horizontalInput, state.offsetHCm);
    set(this.verticalInput, state.offsetVCm);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyProperties(): void {
    const num = (input: FluentTextInput | undefined): number | undefined => {
      const v = Number(input?.value);
      return Number.isFinite(v) && v >= 0 ? Math.round(v * 100) / 100 : undefined;
    };
    const patch: DrawingPropertiesPatch = {
      widthCm: num(this.widthInput) ?? 0,
      heightCm: num(this.heightInput) ?? 0,
      rotationDeg: Number(this.rotationInput?.value) || 0,
      offsetHCm: num(this.horizontalInput) ?? 0,
      offsetVCm: num(this.verticalInput) ?? 0,
    };
    this.$emit("drawing-properties:ok", patch);
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("drawingDialog.title", this);
    if (this.sizeHeading) this.sizeHeading.textContent = t("drawingDialog.size", this);
    if (this.widthLabel) this.widthLabel.textContent = t("drawingDialog.width", this);
    if (this.heightLabel) this.heightLabel.textContent = t("drawingDialog.height", this);
    if (this.rotationLabel) this.rotationLabel.textContent = t("drawingDialog.rotation", this);
    if (this.positionHeading) this.positionHeading.textContent = t("drawingDialog.position", this);
    if (this.horizontalLabel)
      this.horizontalLabel.textContent = t("drawingDialog.horizontal", this);
    if (this.verticalLabel) this.verticalLabel.textContent = t("drawingDialog.vertical", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenDrawingPropertiesDialog;
