import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { TablePropertiesPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";

const CM_TO_TWIPS = 567;

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
  select {
    min-width: 0;
    flex: 1 1 auto;
    box-sizing: border-box;
    font: inherit;
    padding: 3px 4px;
  }
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
  .unit {
    white-space: nowrap;
  }
`;

const template = html<DocenTablePropertiesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="heading" ${ref("sizeHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("preferredWidthLabel")}></label>
          <fluent-text-input ${ref("widthInput")} disabled></fluent-text-input>
          <span class="unit" ${ref("widthUnit")}></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("alignmentLabel")}></label>
          <select ${ref("alignmentSel")}>
            <option value="left"></option>
            <option value="center"></option>
            <option value="right"></option>
          </select>
        </div>
        <div class="field">
          <label ${ref("indentLabel")}></label>
          <fluent-text-input
            ${ref("indentInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("cmUnit")}></span>
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

/**
 * `<docen-table-properties-dialog>` — the Word "Table Properties" dialog's
 * table tab (Table Layout → Properties, or the right-click entry): preferred
 * width (read-only — column geometry is owned by the column grid), alignment
 * and the left indent. The host prefills from the caret table's attrs via
 * `show(attrs)`; OK emits `table-properties:ok` with a
 * {@link TablePropertiesPatch} for the host to stamp via
 * `table-properties-apply`. Rides on `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-table-properties-dialog", template, styles })
class DocenTablePropertiesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable sizeHeading?: HTMLElement;
  @observable preferredWidthLabel?: HTMLElement;
  @observable widthInput?: FluentTextInput;
  @observable widthUnit?: HTMLElement;
  @observable alignmentLabel?: HTMLElement;
  @observable alignmentSel?: HTMLSelectElement;
  @observable indentLabel?: HTMLElement;
  @observable indentInput?: FluentTextInput;
  @observable cmUnit?: HTMLElement;
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

  /** Prefill from the caret table's attrs (verbatim PM mirror of the table
   *  attrs). The indent shows in centimeters (twips → 2-decimal cm). */
  show(attrs: Record<string, unknown> = {}): void {
    if (this.alignmentSel)
      this.alignmentSel.value =
        attrs.alignment === "center" || attrs.alignment === "right" ? attrs.alignment : "left";
    if (this.indentInput) {
      const tw = typeof attrs.indent === "number" ? attrs.indent : 0;
      this.indentInput.value = String(Math.round((tw / CM_TO_TWIPS) * 100) / 100);
    }
    // Preferred width is read-only: the column grid (columnWidths + autofit)
    // owns the geometry — mirroring it here would create a second writer.
    if (this.widthInput) {
      const widths = attrs.columnWidths as number[] | null;
      const total = Array.isArray(widths) ? widths.reduce((a, b) => a + b, 0) : 0;
      this.widthInput.value =
        total > 0 ? String(Math.round((total / CM_TO_TWIPS) * 100) / 100) : "";
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyProperties(): void {
    const cm = Number(this.indentInput?.value);
    const patch: TablePropertiesPatch = {
      alignment: (this.alignmentSel?.value ?? "left") as TablePropertiesPatch["alignment"],
      indent: Number.isFinite(cm) && cm > 0 ? Math.round(cm * CM_TO_TWIPS) : 0,
    };
    this.$emit("table-properties:ok", patch);
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("tableDialog.title", this);
    if (this.sizeHeading) this.sizeHeading.textContent = t("tableDialog.size", this);
    if (this.preferredWidthLabel)
      this.preferredWidthLabel.textContent = t("tableDialog.width", this);
    if (this.widthUnit) this.widthUnit.textContent = t("tableDialog.cm", this);
    if (this.alignmentLabel) this.alignmentLabel.textContent = t("tableDialog.alignment", this);
    if (this.indentLabel) this.indentLabel.textContent = t("tableDialog.indent", this);
    if (this.cmUnit) this.cmUnit.textContent = t("tableDialog.cm", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.alignmentSel) {
      const [left, center, right] = this.alignmentSel.options;
      left.textContent = t("ribbon.cmd.align-left", this);
      center.textContent = t("ribbon.cmd.align-center", this);
      right.textContent = t("ribbon.cmd.align-right", this);
    }
  }
}

export default DocenTablePropertiesDialog;
