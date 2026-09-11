import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import type { DrawingPropertiesPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";
import { measureToCm } from "./measure-input";
// Registration happens through the workspace index's re-export (a type-only
// import here would let esbuild elide the module and drop the side effect).
import type DocenMeasureInput from "./measure-input";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(420px, 92vw);
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
  fluent-dropdown {
    min-width: 0;
    flex: 1 1 auto;
  }
  .unit {
    white-space: nowrap;
  }
  .checks {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px 14px;
  }
`;

/** The position-base vocabularies (ST_RelFromH / ST_RelFromV, the tokens the
 *  Floating attr carries; the dropdown labels localize through
 *  drawingDialog.rel-*). */
const H_BASES = ["column", "margin", "page"] as const;
const V_BASES = ["paragraph", "line", "margin", "page"] as const;

const baseOption = html<string, DocenDrawingPropertiesDialog>`
  <fluent-option value="${(b) => b}"
    >${(b, c) => t(`drawingDialog.rel-${b}`, c.parent)}</fluent-option
  >
`;

const template = html<DocenDrawingPropertiesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="heading" ${ref("sizeHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("widthLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("widthInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("heightLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("heightInput")}></docen-measure-input>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("rotationLabel")}></label>
          <docen-measure-input no-unit ${ref("rotationInput")}></docen-measure-input>
          <span class="unit">°</span>
        </div>
      </div>
      <div class="heading" ${ref("positionHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("horizontalLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("relativeHDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              ${repeat(() => H_BASES, baseOption)}
            </fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
          <docen-measure-input units="cm mm in pt" ${ref("horizontalInput")}></docen-measure-input>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("verticalLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("relativeVDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              ${repeat(() => V_BASES, baseOption)}
            </fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
          <docen-measure-input units="cm mm in pt" ${ref("verticalInput")}></docen-measure-input>
        </div>
      </div>
      <div class="heading" ${ref("layoutHeading")}></div>
      <div class="checks">
        <fluent-field label-position="after">
          <fluent-checkbox slot="input" ${ref("allowOverlapBox")}></fluent-checkbox>
          <label slot="label">${(x) => t("drawingDialog.allow-overlap", x)}</label>
        </fluent-field>
        <fluent-field label-position="after">
          <fluent-checkbox slot="input" ${ref("layoutInCellBox")}></fluent-checkbox>
          <label slot="label">${(x) => t("drawingDialog.layout-in-cell", x)}</label>
        </fluent-field>
        <fluent-field label-position="after">
          <fluent-checkbox slot="input" ${ref("lockAnchorBox")}></fluent-checkbox>
          <label slot="label">${(x) => t("drawingDialog.lock-anchor", x)}</label>
        </fluent-field>
      </div>
      <div class="heading" ${ref("distanceHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("topLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("topInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("bottomLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("bottomInput")}></docen-measure-input>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("leftLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("leftInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("rightLabel")}></label>
          <docen-measure-input units="cm mm in pt" ${ref("rightInput")}></docen-measure-input>
        </div>
      </div>
      <div class="heading" ${ref("altTextHeading")}></div>
      <div class="row">
        <div class="field">
          <fluent-text-input ${ref("altTextInput")}></fluent-text-input>
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

/** The measured boxes (DocenMeasureInput) share the text input's value
 *  surface minus the disabled flag this dialog never sets. */
type MeasureBox = Pick<DocenMeasureInput, "value" | "unit">;

/** The prefill shape the host derives from the selected drawing: everything
 *  already in centimeters (the dialog's display unit), rotation in degrees. */
export interface DrawingPropertiesState {
  widthCm: number;
  heightCm: number;
  rotationDeg: number;
  offsetHCm: number;
  offsetVCm: number;
  /** The position bases (ST_RelFromH/V tokens, drawn from H_BASES/V_BASES). */
  relativeH: string;
  relativeV: string;
  /** The layout flags (Word's Advanced Layout group). */
  allowOverlap: boolean;
  layoutInCell: boolean;
  lockAnchor: boolean;
  /** The wrap distances in cm (Word's Distance from text). */
  distanceCm: { top: number; bottom: number; left: number; right: number };
  /** The replacement text (images only; empty when the drawing has none). */
  altText: string;
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
  @observable widthInput?: MeasureBox;
  @observable heightLabel?: HTMLElement;
  @observable heightInput?: MeasureBox;
  @observable rotationLabel?: HTMLElement;
  @observable rotationInput?: MeasureBox;
  @observable positionHeading?: HTMLElement;
  @observable horizontalLabel?: HTMLElement;
  @observable horizontalInput?: MeasureBox;
  @observable relativeHDropdown?: HTMLElement & { value: string | null };
  @observable verticalLabel?: HTMLElement;
  @observable verticalInput?: MeasureBox;
  @observable relativeVDropdown?: HTMLElement & { value: string | null };
  @observable layoutHeading?: HTMLElement;
  @observable allowOverlapBox?: HTMLElement & { checked: boolean };
  @observable layoutInCellBox?: HTMLElement & { checked: boolean };
  @observable lockAnchorBox?: HTMLElement & { checked: boolean };
  @observable distanceHeading?: HTMLElement;
  @observable topLabel?: HTMLElement;
  @observable topInput?: MeasureBox;
  @observable bottomLabel?: HTMLElement;
  @observable bottomInput?: MeasureBox;
  @observable leftLabel?: HTMLElement;
  @observable leftInput?: MeasureBox;
  @observable rightLabel?: HTMLElement;
  @observable rightInput?: MeasureBox;
  @observable altTextHeading?: HTMLElement;
  @observable altTextInput?: FluentTextInput;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;
  /** The show() prefill at display precision — the OK diff's baseline. */
  #snapshot: DrawingPropertiesState | null = null;

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
   *  cm). Two decimals mirror Word's spin precision. The state snapshots at
   *  that display precision, so OK can drop the fields the user never
   *  touched — writing them back would round-trip the stored value through
   *  the two-decimal display (a 2000-EMU offset would land as 3600). */
  show(state: DrawingPropertiesState): void {
    const round = (v: number): number => Math.round(v * 100) / 100;
    this.#snapshot = {
      ...state,
      widthCm: round(state.widthCm),
      heightCm: round(state.heightCm),
      rotationDeg: round(state.rotationDeg),
      offsetHCm: round(state.offsetHCm),
      offsetVCm: round(state.offsetVCm),
      distanceCm: {
        top: round(state.distanceCm.top),
        bottom: round(state.distanceCm.bottom),
        left: round(state.distanceCm.left),
        right: round(state.distanceCm.right),
      },
    };
    const set = (input: MeasureBox | undefined, cm: number): void => {
      if (input) input.value = round(cm);
    };
    set(this.widthInput, state.widthCm);
    set(this.heightInput, state.heightCm);
    set(this.rotationInput, state.rotationDeg);
    set(this.horizontalInput, state.offsetHCm);
    set(this.verticalInput, state.offsetVCm);
    if (this.relativeHDropdown) this.relativeHDropdown.value = state.relativeH;
    if (this.relativeVDropdown) this.relativeVDropdown.value = state.relativeV;
    if (this.allowOverlapBox) this.allowOverlapBox.checked = state.allowOverlap;
    if (this.layoutInCellBox) this.layoutInCellBox.checked = state.layoutInCell;
    if (this.lockAnchorBox) this.lockAnchorBox.checked = state.lockAnchor;
    set(this.topInput, state.distanceCm.top);
    set(this.bottomInput, state.distanceCm.bottom);
    set(this.leftInput, state.distanceCm.left);
    set(this.rightInput, state.distanceCm.right);
    if (this.altTextInput) this.altTextInput.value = state.altText;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). The position
   *  bases are comboboxes — a typed value outside the token list declines
   *  (the current base survives). Only the fields the user changed ride the
   *  patch; untouched ones keep their stored value un-rounded. */
  applyProperties(): void {
    // A measured box rides its picked unit; the patch stays centimeter-based
    // (the host's writer is untouched), so the value resolves before it lands.
    const num = (input: MeasureBox | undefined): number | undefined => {
      if (input == null) return undefined;
      const v = input.value;
      if (v == null) return undefined;
      const cm = measureToCm(v, input.unit);
      return cm >= 0 ? Math.round(cm * 100) / 100 : undefined;
    };
    const rel = (raw: string | null | undefined, bases: readonly string[]): string | undefined =>
      typeof raw === "string" && (bases as readonly string[]).includes(raw) ? raw : undefined;
    const snap = this.#snapshot;
    const changed = (value: unknown, prefill: unknown): boolean =>
      snap == null || JSON.stringify(value) !== JSON.stringify(prefill);
    const patch: DrawingPropertiesPatch = {};
    const put = <K extends keyof DrawingPropertiesPatch>(
      key: K,
      value: DrawingPropertiesPatch[K],
    ): void => {
      if (value !== undefined && changed(value, snap?.[key])) patch[key] = value;
    };
    put("widthCm", num(this.widthInput));
    put("heightCm", num(this.heightInput));
    const rotationDeg = this.rotationInput?.value;
    if (rotationDeg != null && Number.isFinite(rotationDeg)) put("rotationDeg", rotationDeg);
    put("offsetHCm", num(this.horizontalInput));
    put("offsetVCm", num(this.verticalInput));
    put("relativeH", rel(this.relativeHDropdown?.value, H_BASES));
    put("relativeV", rel(this.relativeVDropdown?.value, V_BASES));
    put("allowOverlap", this.allowOverlapBox?.checked === true);
    put("layoutInCell", this.layoutInCellBox?.checked === true);
    put("lockAnchor", this.lockAnchorBox?.checked === true);
    put("distanceCm", {
      top: num(this.topInput) ?? 0,
      bottom: num(this.bottomInput) ?? 0,
      left: num(this.leftInput) ?? 0,
      right: num(this.rightInput) ?? 0,
    });
    put("altText", this.altTextInput?.value ?? "");
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
    if (this.layoutHeading) this.layoutHeading.textContent = t("drawingDialog.layout", this);
    if (this.distanceHeading) this.distanceHeading.textContent = t("drawingDialog.distance", this);
    if (this.topLabel) this.topLabel.textContent = t("drawingDialog.top", this);
    if (this.bottomLabel) this.bottomLabel.textContent = t("drawingDialog.bottom", this);
    if (this.leftLabel) this.leftLabel.textContent = t("drawingDialog.left", this);
    if (this.rightLabel) this.rightLabel.textContent = t("drawingDialog.right", this);
    if (this.altTextHeading) this.altTextHeading.textContent = t("drawingDialog.altText", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenDrawingPropertiesDialog;
