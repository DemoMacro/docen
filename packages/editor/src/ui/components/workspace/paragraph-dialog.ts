import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { ParagraphDialogPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";

/** Line-spacing select values — multiples encode as w:line 240ths under
 *  "auto"; atLeast/exactly carry points in the value input. */
type LineSpacingChoice = "single" | "lines15" | "double" | "multiple" | "atLeast" | "exactly";

const PT_TO_TWIPS = 20;
const LINE_PER_MULTIPLE = 240;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
  }
  .para-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .para-heading {
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
  .checks {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenParagraphDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="para-body">
      <div class="row">
        <div class="field">
          <label ${ref("alignLabel")}></label>
          <select ${ref("alignSel")}>
            <option value="left"></option>
            <option value="center"></option>
            <option value="right"></option>
            <option value="both"></option>
          </select>
        </div>
        <div class="field">
          <label ${ref("outlineLabel")}></label>
          <select ${ref("outlineSel")}>
            <option value="-1"></option>
            <option value="0"></option>
            <option value="1"></option>
            <option value="2"></option>
            <option value="3"></option>
            <option value="4"></option>
            <option value="5"></option>
            <option value="6"></option>
            <option value="7"></option>
            <option value="8"></option>
          </select>
        </div>
      </div>
      <div class="para-heading" ${ref("indentHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("leftLabel")}></label>
          <fluent-text-input
            ${ref("leftInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("ptA")}></span>
        </div>
        <div class="field">
          <label ${ref("rightLabel")}></label>
          <fluent-text-input
            ${ref("rightInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("ptB")}></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("specialLabel")}></label>
          <select ${ref("specialSel")} @change="${(x) => x.syncSpecialEnabled()}">
            <option value="none"></option>
            <option value="firstLine"></option>
            <option value="hanging"></option>
          </select>
        </div>
        <div class="field">
          <label ${ref("specialValLabel")}></label>
          <fluent-text-input
            ${ref("specialVal")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("ptC")}></span>
        </div>
      </div>
      <div class="para-heading" ${ref("spacingHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("beforeLabel")}></label>
          <fluent-text-input
            ${ref("beforeInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("ptD")}></span>
        </div>
        <div class="field">
          <label ${ref("afterLabel")}></label>
          <fluent-text-input
            ${ref("afterInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
          <span class="unit" ${ref("ptE")}></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("lineLabel")}></label>
          <select ${ref("lineSel")} @change="${(x) => x.syncSpecialEnabled()}">
            <option value="single"></option>
            <option value="lines15"></option>
            <option value="double"></option>
            <option value="multiple"></option>
            <option value="atLeast"></option>
            <option value="exactly"></option>
          </select>
        </div>
        <div class="field">
          <label ${ref("lineValLabel")}></label>
          <fluent-text-input ${ref("lineVal")} type="number" step="any" min="0"></fluent-text-input>
        </div>
      </div>
      <div class="para-heading" ${ref("breaksHeading")}></div>
      <div class="checks">
        <label class="check-field">
          <fluent-checkbox part="widow" ${ref("widow")}></fluent-checkbox>
          <span ${ref("widowLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="keep-next" ${ref("keepNext")}></fluent-checkbox>
          <span ${ref("keepNextLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="keep-lines" ${ref("keepLines")}></fluent-checkbox>
          <span ${ref("keepLinesLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="page-break" ${ref("pageBreak")}></fluent-checkbox>
          <span ${ref("pageBreakLabel")}></span>
        </label>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyParagraph()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** A checkbox widget plus its checked state accessor (fluent-checkbox exposes
 *  `currentChecked`, not the native `checked`). */
type FluentCheckbox = HTMLElement & { currentChecked?: boolean };

/**
 * `<docen-paragraph-dialog>` — the Word "Paragraph" dialog (indent & spacing
 * plus line & page-break controls in one view). The host prefills from the
 * caret paragraph's attrs via `show(attrs)`; OK emits `paragraph:ok` with a
 * full {@link ParagraphDialogPatch} for the host to stamp onto every selected
 * paragraph (`paragraph-dialog-apply`). Cancel / Esc just close. Rides on
 * `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-paragraph-dialog", template, styles })
class DocenParagraphDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable alignLabel?: HTMLElement;
  @observable alignSel?: HTMLSelectElement;
  @observable outlineLabel?: HTMLElement;
  @observable outlineSel?: HTMLSelectElement;
  @observable indentHeading?: HTMLElement;
  @observable leftLabel?: HTMLElement;
  @observable leftInput?: FluentTextInput;
  @observable rightLabel?: HTMLElement;
  @observable rightInput?: FluentTextInput;
  @observable specialLabel?: HTMLElement;
  @observable specialSel?: HTMLSelectElement;
  @observable specialValLabel?: HTMLElement;
  @observable specialVal?: FluentTextInput;
  @observable spacingHeading?: HTMLElement;
  @observable beforeLabel?: HTMLElement;
  @observable beforeInput?: FluentTextInput;
  @observable afterLabel?: HTMLElement;
  @observable afterInput?: FluentTextInput;
  @observable lineLabel?: HTMLElement;
  @observable lineSel?: HTMLSelectElement;
  @observable lineValLabel?: HTMLElement;
  @observable lineVal?: FluentTextInput;
  @observable breaksHeading?: HTMLElement;
  @observable widow?: FluentCheckbox;
  @observable keepNext?: FluentCheckbox;
  @observable keepLines?: FluentCheckbox;
  @observable pageBreak?: FluentCheckbox;
  @observable widowLabel?: HTMLElement;
  @observable keepNextLabel?: HTMLElement;
  @observable keepLinesLabel?: HTMLElement;
  @observable pageBreakLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  @observable ptA?: HTMLElement;
  @observable ptB?: HTMLElement;
  @observable ptC?: HTMLElement;
  @observable ptD?: HTMLElement;
  @observable ptE?: HTMLElement;

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

  /** Prefill every field from the caret paragraph's attrs (verbatim PM
   *  mirror of ParagraphPropertiesOptionsBase). Absent values fall back to
   *  their OOXML defaults (left alignment, single spacing, widow control on). */
  show(attrs: Record<string, unknown> = {}): void {
    const indent = (attrs.indent ?? {}) as {
      left?: number;
      right?: number;
      firstLine?: number;
      hanging?: number;
    };
    const spacing = (attrs.spacing ?? {}) as {
      before?: number;
      after?: number;
      line?: number;
      lineRule?: string;
    };
    if (this.alignSel) this.alignSel.value = (attrs.alignment as string) ?? "left";
    if (this.outlineSel) {
      const level = typeof attrs.outlineLevel === "number" ? attrs.outlineLevel : -1;
      this.outlineSel.value = String(Math.max(-1, Math.min(8, level)));
    }
    if (this.leftInput) this.leftInput.value = this.#pt(indent.left);
    if (this.rightInput) this.rightInput.value = this.#pt(indent.right);
    let special = "none";
    if (indent.firstLine) special = "firstLine";
    else if (indent.hanging) special = "hanging";
    if (this.specialSel) this.specialSel.value = special;
    if (this.specialVal) {
      this.specialVal.value =
        special === "firstLine"
          ? this.#pt(indent.firstLine)
          : special === "hanging"
            ? this.#pt(indent.hanging)
            : "";
    }
    this.syncSpecialEnabled();
    if (this.beforeInput) this.beforeInput.value = this.#pt(spacing.before);
    if (this.afterInput) this.afterInput.value = this.#pt(spacing.after);
    this.#prefillLine(spacing);
    this.#check(this.widow, attrs.widowControl, true);
    this.#check(this.keepNext, attrs.keepNext, false);
    this.#check(this.keepLines, attrs.keepLines, false);
    this.#check(this.pageBreak, attrs.pageBreakBefore, false);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** The value input only applies to multiple/atLeast/exactly spacing. */
  syncSpecialEnabled(): void {
    if (this.lineVal) {
      const choice = (this.lineSel?.value ?? "single") as LineSpacingChoice;
      this.lineVal.disabled = choice === "single" || choice === "lines15" || choice === "double";
    }
    if (this.specialVal && this.specialSel)
      this.specialVal.disabled = this.specialSel.value === "none";
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyParagraph(): void {
    // "Body Text" (-1) commits null — w:outlineLvl has no -1; absence IS
    // body text. Level 0 is a legal level and must survive.
    const outline = this.outlineSel ? Number(this.outlineSel.value) : -1;
    const patch: ParagraphDialogPatch = {
      alignment: this.alignSel?.value ?? "left",
      outlineLevel: outline >= 0 ? outline : null,
      indent: {
        left: this.#twips(this.leftInput?.value),
        right: this.#twips(this.rightInput?.value),
        // firstLine and hanging are mutually exclusive — the unchosen one
        // clears the other (OOXML rejects both on one paragraph).
        firstLine: undefined,
        hanging: undefined,
      },
      spacing: {},
      widowControl: this.widow?.currentChecked ?? true,
      keepNext: this.keepNext?.currentChecked ?? false,
      keepLines: this.keepLines?.currentChecked ?? false,
      pageBreakBefore: this.pageBreak?.currentChecked ?? false,
    };
    const special = this.specialSel?.value;
    if (special === "firstLine") patch.indent.firstLine = this.#twips(this.specialVal?.value);
    else if (special === "hanging") patch.indent.hanging = this.#twips(this.specialVal?.value);
    const choice = (this.lineSel?.value ?? "single") as LineSpacingChoice;
    if (choice === "single") patch.spacing = { line: 240, lineRule: "auto" };
    else if (choice === "lines15") patch.spacing = { line: 360, lineRule: "auto" };
    else if (choice === "double") patch.spacing = { line: 480, lineRule: "auto" };
    else if (choice === "multiple")
      patch.spacing = {
        line: Math.round((this.#num(this.lineVal?.value) || 1) * LINE_PER_MULTIPLE),
        lineRule: "auto",
      };
    else if (choice === "atLeast")
      patch.spacing = {
        line: Math.round(this.#num(this.lineVal?.value) * PT_TO_TWIPS),
        lineRule: "atLeast",
      };
    else
      patch.spacing = {
        line: Math.round(this.#num(this.lineVal?.value) * PT_TO_TWIPS),
        lineRule: "exact",
      };
    patch.spacing.before = this.#twips(this.beforeInput?.value);
    patch.spacing.after = this.#twips(this.afterInput?.value);
    this.$emit("paragraph:ok", patch);
    this.hide();
  }

  /** Twips → points, rounded to 2 decimals for the input. */
  #pt(twips?: number): string {
    if (typeof twips !== "number") return "";
    return String(Math.round((twips / PT_TO_TWIPS) * 100) / 100);
  }

  #twips(pt?: string): number | undefined {
    if (pt === undefined || pt === "") return undefined;
    const n = Number(pt);
    return Number.isFinite(n) ? Math.round(n * PT_TO_TWIPS) : undefined;
  }

  #num(v?: string): number {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  #prefillLine(spacing: { line?: number; lineRule?: string }): void {
    if (!this.lineSel) return;
    const line = spacing.line ?? LINE_PER_MULTIPLE;
    const rule = spacing.lineRule ?? "auto";
    let choice: LineSpacingChoice = "single";
    let val = "";
    if (rule === "atLeast" || rule === "exact") {
      // OOXML's token is "exact"; the UI choice spells it "exactly".
      choice = rule === "atLeast" ? "atLeast" : "exactly";
      val = this.#pt(line);
    } else {
      const mult = line / LINE_PER_MULTIPLE;
      if (Math.abs(mult - 1) < 0.01) choice = "single";
      else if (Math.abs(mult - 1.5) < 0.01) choice = "lines15";
      else if (Math.abs(mult - 2) < 0.01) choice = "double";
      else {
        choice = "multiple";
        val = String(Math.round(mult * 100) / 100);
      }
    }
    this.lineSel.value = choice;
    if (this.lineVal) this.lineVal.value = val;
    this.syncSpecialEnabled();
  }

  #check(box: FluentCheckbox | undefined, value: unknown, fallback: boolean): void {
    if (box) box.currentChecked = typeof value === "boolean" ? value : fallback;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("paragraph.title", this);
    if (this.alignLabel) this.alignLabel.textContent = t("paragraph.alignment", this);
    if (this.outlineLabel) this.outlineLabel.textContent = t("paragraph.outline", this);
    if (this.indentHeading) this.indentHeading.textContent = t("paragraph.indentHeading", this);
    if (this.leftLabel) this.leftLabel.textContent = t("paragraph.left", this);
    if (this.rightLabel) this.rightLabel.textContent = t("paragraph.right", this);
    if (this.specialLabel) this.specialLabel.textContent = t("paragraph.special", this);
    if (this.specialValLabel) this.specialValLabel.textContent = t("paragraph.value", this);
    if (this.spacingHeading) this.spacingHeading.textContent = t("paragraph.spacingHeading", this);
    if (this.beforeLabel) this.beforeLabel.textContent = t("paragraph.before", this);
    if (this.afterLabel) this.afterLabel.textContent = t("paragraph.after", this);
    if (this.lineLabel) this.lineLabel.textContent = t("paragraph.lineSpacing", this);
    if (this.lineValLabel) this.lineValLabel.textContent = t("paragraph.setValue", this);
    if (this.breaksHeading) this.breaksHeading.textContent = t("paragraph.breaksHeading", this);
    if (this.widowLabel) this.widowLabel.textContent = t("paragraph.widowControl", this);
    if (this.keepNextLabel) this.keepNextLabel.textContent = t("paragraph.keepNext", this);
    if (this.keepLinesLabel) this.keepLinesLabel.textContent = t("paragraph.keepLines", this);
    if (this.pageBreakLabel) this.pageBreakLabel.textContent = t("paragraph.pageBreakBefore", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    // The unit chips (磅) after each measurement input.
    for (const el of [this.ptA, this.ptB, this.ptC, this.ptD, this.ptE])
      if (el) el.textContent = t("paragraph.pt", this);
    // Select option labels (alignment reuses the ribbon's entries).
    if (this.alignSel) {
      const [left, center, right, both] = this.alignSel.options;
      left.textContent = t("ribbon.cmd.align-left", this);
      center.textContent = t("ribbon.cmd.align-center", this);
      right.textContent = t("ribbon.cmd.align-right", this);
      both.textContent = t("ribbon.cmd.justify", this);
    }
    if (this.outlineSel) {
      const opts = this.outlineSel.options;
      opts[0].textContent = t("paragraph.outlineBody", this);
      for (let i = 1; i < opts.length; i++)
        opts[i].textContent = `${t("paragraph.level", this)} ${i}`;
    }
    if (this.specialSel) {
      const [none, firstLine, hanging] = this.specialSel.options;
      none.textContent = t("paragraph.specialNone", this);
      firstLine.textContent = t("paragraph.specialFirstLine", this);
      hanging.textContent = t("paragraph.specialHanging", this);
    }
    if (this.lineSel) {
      const labels = this.lineSel.options;
      labels[0].textContent = t("paragraph.lsSingle", this);
      labels[1].textContent = t("paragraph.ls15", this);
      labels[2].textContent = t("paragraph.lsDouble", this);
      labels[3].textContent = t("paragraph.lsMultiple", this);
      labels[4].textContent = t("paragraph.lsAtLeast", this);
      labels[5].textContent = t("paragraph.lsExactly", this);
    }
  }
}

export default DocenParagraphDialog;
