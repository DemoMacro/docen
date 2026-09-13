import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string };

/** A `fluent-dropdown` combobox: `value` reads/writes the picked option. */
type FluentDropdown = HTMLElement & { value: string };

/** One note kind's document-level numbering settings — the dialog's fields
 *  map 1:1 onto the settings.xml w:footnotePr/w:endnotePr children. */
export interface NoteKindSettings {
  /** Note placement (w:pos) — footnotes use pageBottom/beneathText, endnotes
   *  sectEnd/docEnd (each dropdown offers only its own two). */
  pos: "pageBottom" | "beneathText" | "sectEnd" | "docEnd";
  /** Number format token (w:numFmt) — one of the dialog's closed list. */
  numFmt: string;
  /** Starting number (w:numStart), ≥ 1. */
  numStart: number;
  /** When numbering restarts (w:numRestart). */
  numRestart: "continuous" | "eachSect" | "eachPage";
}

/** The dialog's two committed groups. */
export interface NoteSettingsValues {
  footnote: NoteKindSettings;
  endnote: NoteKindSettings;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(400px, 92vw);
  }
  .note-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .note-heading {
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
`;

const template = html<DocenNoteSettingsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="note-body">
      <div class="note-heading" ${ref("footnoteHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("footnotePosLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("footnotePosDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="pageBottom"></fluent-option>
              <fluent-option value="beneathText"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
        <div class="field">
          <label ${ref("footnoteFmtLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("footnoteFmtDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="decimal">1, 2, 3, …</fluent-option>
              <fluent-option value="lowerLetter">a, b, c, …</fluent-option>
              <fluent-option value="upperLetter">A, B, C, …</fluent-option>
              <fluent-option value="lowerRoman">i, ii, iii, …</fluent-option>
              <fluent-option value="upperRoman">I, II, III, …</fluent-option>
              <fluent-option value="chineseCountingThousand">一, 二, 三, …</fluent-option>
              <fluent-option value="chineseLegalSimplified">壹, 贰, 叁, …</fluent-option>
              <fluent-option value="ideographTraditional">甲, 乙, 丙, …</fluent-option>
              <fluent-option value="ideographZodiac">子, 丑, 寅, …</fluent-option>
              <fluent-option value="decimalEnclosedCircle">①, ②, ③, …</fluent-option>
              <fluent-option value="aiueo">あ, い, う, …</fluent-option>
              <fluent-option value="aiueoFullWidth">ア, イ, ウ, …</fluent-option>
              <fluent-option value="iroha">い, ろ, は, …</fluent-option>
              <fluent-option value="irohaFullWidth">イ, ロ, ハ, …</fluent-option>
              <fluent-option value="ganada">가, 나, 다, …</fluent-option>
              <fluent-option value="chosung">ㄱ, ㄴ, ㄷ, …</fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("footnoteStartLabel")}></label>
          <fluent-text-input
            ${ref("footnoteStartInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("footnoteRestartLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("footnoteRestartDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="continuous"></fluent-option>
              <fluent-option value="eachSect"></fluent-option>
              <fluent-option value="eachPage"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="note-heading" ${ref("endnoteHeading")}></div>
      <div class="row">
        <div class="field">
          <label ${ref("endnotePosLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("endnotePosDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="sectEnd"></fluent-option>
              <fluent-option value="docEnd"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
        <div class="field">
          <label ${ref("endnoteFmtLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("endnoteFmtDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="decimal">1, 2, 3, …</fluent-option>
              <fluent-option value="lowerLetter">a, b, c, …</fluent-option>
              <fluent-option value="upperLetter">A, B, C, …</fluent-option>
              <fluent-option value="lowerRoman">i, ii, iii, …</fluent-option>
              <fluent-option value="upperRoman">I, II, III, …</fluent-option>
              <fluent-option value="chineseCountingThousand">一, 二, 三, …</fluent-option>
              <fluent-option value="chineseLegalSimplified">壹, 贰, 叁, …</fluent-option>
              <fluent-option value="ideographTraditional">甲, 乙, 丙, …</fluent-option>
              <fluent-option value="ideographZodiac">子, 丑, 寅, …</fluent-option>
              <fluent-option value="decimalEnclosedCircle">①, ②, ③, …</fluent-option>
              <fluent-option value="aiueo">あ, い, う, …</fluent-option>
              <fluent-option value="aiueoFullWidth">ア, イ, ウ, …</fluent-option>
              <fluent-option value="iroha">い, ろ, は, …</fluent-option>
              <fluent-option value="irohaFullWidth">イ, ロ, ハ, …</fluent-option>
              <fluent-option value="ganada">가, 나, 다, …</fluent-option>
              <fluent-option value="chosung">ㄱ, ㄴ, ㄷ, …</fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("endnoteStartLabel")}></label>
          <fluent-text-input
            ${ref("endnoteStartInput")}
            type="number"
            step="1"
            min="1"
          ></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("endnoteRestartLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("endnoteRestartDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="continuous"></fluent-option>
              <fluent-option value="eachSect"></fluent-option>
              <fluent-option value="eachPage"></fluent-option>
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
        @click="${(x) => x.applySettings()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-note-settings-dialog>` — Word's "Footnote and Endnote" dialog
 * (References tab's footnotes-group launcher): per-kind placement, number
 * format, start number, and restart mode. Both kinds show at once (Word
 * radio-switches them; the two sections are the same settings without the
 * radio state). Opened via `show(values)` from documentExtras.settings and
 * committed via `note-settings:ok`.
 */
@customElement({ name: "docen-note-settings-dialog", template, styles })
class DocenNoteSettingsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable footnoteHeading?: HTMLElement;
  @observable footnotePosLabel?: HTMLElement;
  @observable footnoteFmtLabel?: HTMLElement;
  @observable footnoteStartLabel?: HTMLElement;
  @observable footnoteRestartLabel?: HTMLElement;
  @observable footnotePosDropdown?: FluentDropdown;
  @observable footnoteFmtDropdown?: FluentDropdown;
  @observable footnoteStartInput?: FluentTextInput;
  @observable footnoteRestartDropdown?: FluentDropdown;
  @observable endnoteHeading?: HTMLElement;
  @observable endnotePosLabel?: HTMLElement;
  @observable endnoteFmtLabel?: HTMLElement;
  @observable endnoteStartLabel?: HTMLElement;
  @observable endnoteRestartLabel?: HTMLElement;
  @observable endnotePosDropdown?: FluentDropdown;
  @observable endnoteFmtDropdown?: FluentDropdown;
  @observable endnoteStartInput?: FluentTextInput;
  @observable endnoteRestartDropdown?: FluentDropdown;
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

  /** Prefill from documentExtras.settings; absent fields fall back to Word's
   *  footnote (1,2,3 at page bottom) and endnote (i,ii,iii at section end)
   *  defaults. */
  show(
    values: {
      footnote?: Partial<NoteKindSettings>;
      endnote?: Partial<NoteKindSettings>;
    } = {},
  ): void {
    const fn = values.footnote ?? {};
    const en = values.endnote ?? {};
    if (this.footnotePosDropdown)
      this.footnotePosDropdown.value = fn.pos === "beneathText" ? "beneathText" : "pageBottom";
    if (this.endnotePosDropdown)
      this.endnotePosDropdown.value = en.pos === "docEnd" ? "docEnd" : "sectEnd";
    if (this.footnoteFmtDropdown) this.footnoteFmtDropdown.value = fn.numFmt ?? "decimal";
    if (this.endnoteFmtDropdown) this.endnoteFmtDropdown.value = en.numFmt ?? "lowerRoman";
    if (this.footnoteStartInput)
      this.footnoteStartInput.value = String(typeof fn.numStart === "number" ? fn.numStart : 1);
    if (this.endnoteStartInput)
      this.endnoteStartInput.value = String(typeof en.numStart === "number" ? en.numStart : 1);
    if (this.footnoteRestartDropdown)
      this.footnoteRestartDropdown.value = fn.numRestart ?? "continuous";
    if (this.endnoteRestartDropdown)
      this.endnoteRestartDropdown.value = en.numRestart ?? "continuous";
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applySettings(): void {
    this.$emit("note-settings:ok", {
      footnote: {
        pos: (this.footnotePosDropdown?.value ?? "pageBottom") as NoteKindSettings["pos"],
        // The dropdown only offers its own tokens — the value reads back
        // closed, no wider validation needed.
        numFmt: this.footnoteFmtDropdown?.value ?? "decimal",
        numStart: Math.max(1, Math.round(Number(this.footnoteStartInput?.value) || 1)),
        numRestart: (this.footnoteRestartDropdown?.value ??
          "continuous") as NoteKindSettings["numRestart"],
      },
      endnote: {
        pos: (this.endnotePosDropdown?.value ?? "sectEnd") as NoteKindSettings["pos"],
        numFmt: this.endnoteFmtDropdown?.value ?? "lowerRoman",
        numStart: Math.max(1, Math.round(Number(this.endnoteStartInput?.value) || 1)),
        numRestart: (this.endnoteRestartDropdown?.value ??
          "continuous") as NoteKindSettings["numRestart"],
      },
    } satisfies NoteSettingsValues);
    this.hide();
  }

  #labelOptions(dropdown: FluentDropdown | undefined, labels: string[]): void {
    if (!dropdown) return;
    dropdown.querySelectorAll("fluent-option").forEach((opt, i) => {
      if (labels[i]) opt.textContent = labels[i];
    });
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("note.settings-title", this);
    if (this.footnoteHeading) this.footnoteHeading.textContent = t("note.footnote-title", this);
    if (this.endnoteHeading) this.endnoteHeading.textContent = t("note.endnote-title", this);
    if (this.footnotePosLabel) this.footnotePosLabel.textContent = t("note.pos", this);
    if (this.endnotePosLabel) this.endnotePosLabel.textContent = t("note.pos", this);
    this.#labelOptions(this.footnotePosDropdown, [
      t("note.pos-page-bottom", this),
      t("note.pos-beneath-text", this),
    ]);
    this.#labelOptions(this.endnotePosDropdown, [
      t("note.pos-sect-end", this),
      t("note.pos-doc-end", this),
    ]);
    if (this.footnoteFmtLabel) this.footnoteFmtLabel.textContent = t("note.num-fmt", this);
    if (this.endnoteFmtLabel) this.endnoteFmtLabel.textContent = t("note.num-fmt", this);
    if (this.footnoteStartLabel) this.footnoteStartLabel.textContent = t("note.num-start", this);
    if (this.endnoteStartLabel) this.endnoteStartLabel.textContent = t("note.num-start", this);
    if (this.footnoteRestartLabel)
      this.footnoteRestartLabel.textContent = t("note.num-restart", this);
    if (this.endnoteRestartLabel)
      this.endnoteRestartLabel.textContent = t("note.num-restart", this);
    this.#labelOptions(this.footnoteRestartDropdown, [
      t("note.restart-continuous", this),
      t("note.restart-each-sect", this),
      t("note.restart-each-page", this),
    ]);
    this.#labelOptions(this.endnoteRestartDropdown, [
      t("note.restart-continuous", this),
      t("note.restart-each-sect", this),
      t("note.restart-each-page", this),
    ]);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenNoteSettingsDialog;
