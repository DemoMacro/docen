import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** A `fluent-dropdown` combobox: `value` reads/writes the picked option. */
type FluentDropdown = HTMLElement & { value: string };

/** The dialog's values in centimeters — the unit Word's zh dialogs show. All
 *  conversion to/from OOXML twips stays on the host. */
export interface PageSetupValues {
  margins: { top: number; bottom: number; left: number; right: number };
  size: { width: number; height: number };
  /** Section vertical alignment (w:vAlign). */
  verticalAlign: "top" | "center" | "both" | "bottom";
  /** Binding gutter width in cm (w:pgMar gutter) — 0 = none. */
  gutter?: number;
  /** Header/footer distance from the page edge in cm (w:pgMar header/footer). */
  headerDistance?: number;
  footerDistance?: number;
  /** First-page-different header/footer (w:titlePg). */
  titlePage?: boolean;
  /** Section start (w:type) — "nextPage" is Word's omitted default. */
  sectionStart?: "nextPage" | "continuous" | "oddPage" | "evenPage";
  /** Document grid (w:docGrid): behavior + the per-line/per-page counts; the
   *  lines ↔ linePitch and chars ↔ charSpace conversions need the page's
   *  usable box and the Normal font pitch, so they stay on the host. */
  grid?: {
    type?: "default" | "lines" | "linesAndChars" | "snapToChars";
    charsPerLine?: number;
    linesPerPage?: number;
  };
}

/** Word's defaults (Normal margins on A4) for absent prefill fields. */
const DEFAULTS = {
  margin: 2.54,
  side: 3.18,
  width: 21,
  height: 29.7,
  header: 1.5,
  footer: 1.75,
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
  .unit,
  .count-unit {
    white-space: nowrap;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
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
      <div class="row">
        <div class="field">
          <label ${ref("gutterLabel")}></label>
          <fluent-text-input
            ${ref("gutterInput")}
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
      <div class="row">
        <div class="field">
          <label ${ref("verticalAlignLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("verticalAlignDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="top"></fluent-option>
              <fluent-option value="center"></fluent-option>
              <fluent-option value="both"></fluent-option>
              <fluent-option value="bottom"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="setup-heading" ${ref("layoutHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("sectionStartLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("sectionStartDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="nextPage"></fluent-option>
              <fluent-option value="continuous"></fluent-option>
              <fluent-option value="oddPage"></fluent-option>
              <fluent-option value="evenPage"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("headerLabel")}></label>
          <fluent-text-input
            ${ref("headerInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
        <div class="field">
          <label ${ref("footerLabel")}></label>
          <fluent-text-input
            ${ref("footerInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit"></span>
        </div>
      </div>
      <!-- fluent-checkbox has no default label slot (indicator slots only) —
           the label span sits outside, the wrapping <label> routes clicks. -->
      <label class="check-field">
        <fluent-checkbox ${ref("titlePageBox")}></fluent-checkbox>
        <span ${ref("titlePageLabel")}></span>
      </label>
      <div class="setup-heading" ${ref("gridHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("gridTypeLabel")}></label>
          <fluent-dropdown
            type="combobox"
            appearance="outline"
            ${ref("gridTypeDropdown")}
            @change="${(x) => x.onGridTypeChange()}"
          >
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="default"></fluent-option>
              <fluent-option value="lines"></fluent-option>
              <fluent-option value="linesAndChars"></fluent-option>
              <fluent-option value="snapToChars"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("charsLabel")}></label>
          <fluent-text-input
            ${ref("charsInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
          <span class="count-unit" ${ref("charsUnit")}></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("linesLabel")}></label>
          <fluent-text-input
            ${ref("linesInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
          <span class="count-unit" ${ref("linesUnit")}></span>
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
 * `<docen-page-setup-dialog>` — the Word "Page Setup" geometry fields: margins
 * + gutter, paper size, vertical alignment, the layout group (section start,
 * header/footer distances, first-page-different), and the document grid
 * (behavior + the per-line/per-page counts) — measurements in centimeters.
 * Opened by the Margins
 * menu's Custom Margins and the Size menu's More Paper Sizes items; the host
 * prefills from the current section via `show(values)` and commits via
 * `page-setup:ok`. Rides on `<docen-dialog>` for the modal shell.
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
  @observable gutterLabel?: HTMLElement;
  @observable gutterInput?: FluentTextInput;
  @observable sizeHeading?: HTMLElement;
  @observable widthLabel?: HTMLElement;
  @observable widthInput?: FluentTextInput;
  @observable heightLabel?: HTMLElement;
  @observable heightInput?: FluentTextInput;
  @observable verticalAlignLabel?: HTMLElement;
  @observable verticalAlignDropdown?: FluentDropdown;
  @observable layoutHeading?: HTMLElement;
  @observable sectionStartLabel?: HTMLElement;
  @observable sectionStartDropdown?: FluentDropdown;
  @observable headerLabel?: HTMLElement;
  @observable headerInput?: FluentTextInput;
  @observable footerLabel?: HTMLElement;
  @observable footerInput?: FluentTextInput;
  @observable titlePageBox?: HTMLElement & { checked?: boolean };
  @observable titlePageLabel?: HTMLElement;
  @observable gridHeading?: HTMLElement;
  @observable gridTypeLabel?: HTMLElement;
  @observable gridTypeDropdown?: FluentDropdown;
  @observable charsLabel?: HTMLElement;
  @observable charsInput?: FluentTextInput;
  @observable charsUnit?: HTMLElement;
  @observable linesLabel?: HTMLElement;
  @observable linesInput?: FluentTextInput;
  @observable linesUnit?: HTMLElement;
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
      verticalAlign?: PageSetupValues["verticalAlign"];
      gutter?: number;
      headerDistance?: number;
      footerDistance?: number;
      titlePage?: boolean;
      sectionStart?: NonNullable<PageSetupValues["sectionStart"]>;
      grid?: NonNullable<PageSetupValues["grid"]>;
    } = {},
  ): void {
    const margins = values.margins ?? {};
    const size = values.size ?? {};
    if (this.topInput) this.topInput.value = this.#cm(margins.top, DEFAULTS.margin);
    if (this.bottomInput) this.bottomInput.value = this.#cm(margins.bottom, DEFAULTS.margin);
    if (this.leftInput) this.leftInput.value = this.#cm(margins.left, DEFAULTS.side);
    if (this.rightInput) this.rightInput.value = this.#cm(margins.right, DEFAULTS.side);
    if (this.gutterInput) this.gutterInput.value = this.#cm(values.gutter ?? 0, 0);
    if (this.widthInput) this.widthInput.value = this.#cm(size.width, DEFAULTS.width);
    if (this.heightInput) this.heightInput.value = this.#cm(size.height, DEFAULTS.height);
    if (this.verticalAlignDropdown)
      this.verticalAlignDropdown.value = values.verticalAlign ?? "top";
    if (this.sectionStartDropdown)
      this.sectionStartDropdown.value = values.sectionStart ?? "nextPage";
    if (this.headerInput) this.headerInput.value = this.#cm(values.headerDistance, DEFAULTS.header);
    if (this.footerInput) this.footerInput.value = this.#cm(values.footerDistance, DEFAULTS.footer);
    if (this.titlePageBox) this.titlePageBox.checked = values.titlePage === true;
    const gridType = values.grid?.type ?? "lines";
    if (this.gridTypeDropdown) this.gridTypeDropdown.value = gridType;
    // "No grid" grays both counts; a lines-only grid has no per-line count
    // (the char grid is what defines "per line").
    const charsOk = gridType === "linesAndChars" || gridType === "snapToChars";
    if (this.charsInput) {
      const chars = values.grid?.charsPerLine;
      this.charsInput.value =
        typeof chars === "number" && chars > 0 ? String(Math.round(chars)) : "";
      this.charsInput.disabled = !charsOk;
    }
    if (this.linesInput) {
      this.linesInput.value =
        typeof values.grid?.linesPerPage === "number" && values.grid.linesPerPage > 0
          ? String(Math.round(values.grid.linesPerPage))
          : "";
      this.linesInput.disabled = gridType === "default";
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Mirror Word's enablement: "no grid" grays both counts; a lines-only
   *  grid grays the per-line count. */
  readonly onGridTypeChange = (): void => {
    const gridType = this.gridTypeDropdown?.value;
    if (this.charsInput)
      this.charsInput.disabled = gridType !== "linesAndChars" && gridType !== "snapToChars";
    if (this.linesInput) this.linesInput.disabled = gridType === "default";
  };

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
    const vAlign = String(this.verticalAlignDropdown?.value ?? "top");
    const verticalAlign = (["top", "center", "both", "bottom"] as const).includes(
      vAlign as PageSetupValues["verticalAlign"],
    )
      ? (vAlign as PageSetupValues["verticalAlign"])
      : "top";
    const gutter = this.#nonNeg(this.gutterInput?.value);
    // The dropdowns only offer their own tokens — the value reads back
    // closed, no wider validation needed.
    const sectionStart = (this.sectionStartDropdown?.value ?? "nextPage") as NonNullable<
      PageSetupValues["sectionStart"]
    >;
    const gridType = (this.gridTypeDropdown?.value ?? "lines") as NonNullable<
      PageSetupValues["grid"]
    >["type"];
    // A cleared count field keeps the document's current pitch/charSpace
    // (undefined) rather than committing 0. A grayed per-line field keeps
    // its stale value — the host only commits chars when the grid uses them.
    const chars = Number(this.charsInput?.value);
    const lines = Number(this.linesInput?.value);
    this.$emit("page-setup:ok", {
      margins,
      size,
      verticalAlign,
      gutter,
      headerDistance: this.#num(this.headerInput?.value, DEFAULTS.header),
      footerDistance: this.#num(this.footerInput?.value, DEFAULTS.footer),
      titlePage: this.titlePageBox?.checked === true,
      sectionStart,
      grid: {
        type: gridType,
        charsPerLine: Number.isFinite(chars) && chars > 0 ? Math.round(chars) : undefined,
        linesPerPage: Number.isFinite(lines) && lines > 0 ? Math.round(lines) : undefined,
      },
    } satisfies PageSetupValues);
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

  /** Gutter accepts 0 (no binding allowance); a cleared field stays 0. */
  #nonNeg(v: string | undefined): number {
    const n = Number(v);
    return Number.isFinite(n) && n >= 0 ? n : 0;
  }

  #labelOptions(dropdown: FluentDropdown | undefined, labels: string[]): void {
    if (!dropdown) return;
    dropdown.querySelectorAll("fluent-option").forEach((opt, i) => {
      if (labels[i]) opt.textContent = labels[i];
    });
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("ribbon.group.page-setup", this);
    if (this.marginsHeading) this.marginsHeading.textContent = t("ribbon.cmd.margins", this);
    if (this.sizeHeading) this.sizeHeading.textContent = t("ribbon.cmd.page-size", this);
    if (this.topLabel) this.topLabel.textContent = t("pageSetup.top", this);
    if (this.bottomLabel) this.bottomLabel.textContent = t("pageSetup.bottom", this);
    if (this.leftLabel) this.leftLabel.textContent = t("pageSetup.left", this);
    if (this.rightLabel) this.rightLabel.textContent = t("pageSetup.right", this);
    if (this.gutterLabel) this.gutterLabel.textContent = t("pageSetup.gutter", this);
    if (this.widthLabel) this.widthLabel.textContent = t("pageSetup.width", this);
    if (this.heightLabel) this.heightLabel.textContent = t("pageSetup.height", this);
    if (this.verticalAlignLabel)
      this.verticalAlignLabel.textContent = t("pageSetup.verticalAlign", this);
    this.#labelOptions(this.verticalAlignDropdown, [
      t("pageSetup.vAlignTop", this),
      t("pageSetup.vAlignCenter", this),
      t("pageSetup.vAlignBoth", this),
      t("pageSetup.vAlignBottom", this),
    ]);
    if (this.layoutHeading) this.layoutHeading.textContent = t("pageSetup.layout", this);
    if (this.sectionStartLabel)
      this.sectionStartLabel.textContent = t("pageSetup.sectionStart", this);
    this.#labelOptions(this.sectionStartDropdown, [
      t("pageSetup.startNextPage", this),
      t("pageSetup.startContinuous", this),
      t("pageSetup.startOddPage", this),
      t("pageSetup.startEvenPage", this),
    ]);
    if (this.headerLabel) this.headerLabel.textContent = t("pageSetup.header", this);
    if (this.footerLabel) this.footerLabel.textContent = t("pageSetup.footer", this);
    if (this.titlePageLabel) this.titlePageLabel.textContent = t("pageSetup.titlePage", this);
    if (this.gridHeading) this.gridHeading.textContent = t("pageSetup.grid", this);
    if (this.gridTypeLabel) this.gridTypeLabel.textContent = t("pageSetup.gridType", this);
    this.#labelOptions(this.gridTypeDropdown, [
      t("pageSetup.gridNone", this),
      t("pageSetup.gridLines", this),
      t("pageSetup.gridLinesAndChars", this),
      t("pageSetup.gridSnapToChars", this),
    ]);
    if (this.charsLabel) this.charsLabel.textContent = t("pageSetup.charsPerLine", this);
    if (this.charsUnit) this.charsUnit.textContent = t("pageSetup.chars", this);
    if (this.linesLabel) this.linesLabel.textContent = t("pageSetup.linesPerPage", this);
    if (this.linesUnit) this.linesUnit.textContent = t("pageSetup.lines", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    // The cm unit chips after each input share one text; the grid count
    // units are stamped via their refs (they are not cm).
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".unit") ?? [])
      el.textContent = t("pageSetup.cm", this);
  }
}

export default DocenPageSetupDialog;
