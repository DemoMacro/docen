import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string; disabled: boolean };

/** A `fluent-dropdown` combobox: `value` reads/writes the picked option. */
type FluentDropdown = HTMLElement & { value: string };

type FluentRadioGroup = HTMLElement & { value: string };

/** The Page Number Format dialog's values — the w:numFmt token, and whether
 *  this section restarts the count (continue-from-previous vs start-at). */
export interface PageNumberFormatValues {
  /** The w:numFmt token (the dropdown only offers its own tokens). */
  format: string;
  /** Continue the previous section's numbering (no w:start written). */
  continueFromPrevious: boolean;
  /** The restart number (w:start) when not continuing. */
  start: number;
}

/** The 编号格式 options — Word's page-number format list, each entry showing
 *  its rendered run; the value is the OOXML token verbatim. */
const FORMATS: [string, string][] = [
  ["decimal", "1, 2, 3, …"],
  ["lowerLetter", "a, b, c, …"],
  ["upperLetter", "A, B, C, …"],
  ["lowerRoman", "i, ii, iii, …"],
  ["upperRoman", "I, II, III, …"],
  ["decimalEnclosedCircle", "①, ②, ③, …"],
  ["chineseCounting", "一, 二, 三, …"],
  ["chineseLegalSimplified", "壹, 贰, 叁, …"],
  ["aiueo", "あ, い, う, …"],
  ["iroha", "い, ろ, は, …"],
  ["katakana", "ア, イ, ウ, …"],
  ["katakanaHalf", "ｱ, ｲ, ｳ, …"],
  ["koreanCounting", "가, 나, 다, …"],
  ["ganada", "ㄱ, ㄴ, ㄷ, …"],
  ["numberInDash", "- 1 -, - 2 -, …"],
];

const formatOptionTemplate = html<(typeof FORMATS)[number], DocenPageNumberFormatDialog>`
  <fluent-option value="${(f) => f[0]}">${(f) => f[1]}</fluent-option>
`;

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
  .row > label {
    white-space: nowrap;
  }
  fluent-dropdown {
    flex: 1 1 auto;
    min-width: 0;
  }
  fluent-radio-group {
    gap: 6px;
  }
  fluent-text-input {
    width: 90px;
  }
`;

const template = html<DocenPageNumberFormatDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <label ${ref("formatLabel")}></label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("formatDropdown")}>
          <fluent-listbox popover="manual" tabindex="-1">
            ${repeat(() => FORMATS, formatOptionTemplate)}
          </fluent-listbox>
          <input slot="control" role="combobox" aria-readonly="true" readonly />
        </fluent-dropdown>
      </div>
      <div class="row">
        <label ${ref("numberingLabel")}></label>
      </div>
      <fluent-radio-group
        orientation="vertical"
        ${ref("modeGroup")}
        @change="${(x) => x.syncMode()}"
      >
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="continue"></fluent-radio>
          <label slot="label" ${ref("continueLabel")}></label>
        </fluent-field>
        <fluent-field label-position="after">
          <fluent-radio slot="input" value="start"></fluent-radio>
          <label slot="label">
            <span ${ref("startLabel")}></span>
            <fluent-text-input
              ${ref("startInput")}
              type="number"
              step="1"
              min="0"
            ></fluent-text-input>
          </label>
        </fluent-field>
      </fluent-radio-group>
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
 * `<docen-page-number-format-dialog>` — Word's "Page Number Format" fields
 * (the Page Number menu's format entry): the numbering format token and the
 * restart mode (continue the previous section vs a start value). Opened via
 * `show(values)` from the current section's w:pgNumType and committed via
 * `page-number-format:ok`.
 */
@customElement({ name: "docen-page-number-format-dialog", template, styles })
class DocenPageNumberFormatDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable formatLabel?: HTMLElement;
  @observable formatDropdown?: FluentDropdown;
  @observable numberingLabel?: HTMLElement;
  @observable modeGroup?: FluentRadioGroup;
  @observable continueLabel?: HTMLElement;
  @observable startLabel?: HTMLElement;
  @observable startInput?: FluentTextInput;
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

  /** Prefill from the current section's w:pgNumType; absent fields fall back
   *  to Word's defaults (decimal, continuing). */
  show(values: Partial<PageNumberFormatValues> = {}): void {
    if (this.formatDropdown) this.formatDropdown.value = values.format ?? "decimal";
    if (this.modeGroup) this.modeGroup.value = values.continueFromPrevious ? "continue" : "start";
    if (this.startInput) this.startInput.value = String(values.start ?? 1);
    this.syncMode();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyOptions(): void {
    this.$emit("page-number-format:ok", {
      format: this.formatDropdown?.value ?? "decimal",
      continueFromPrevious: (this.modeGroup?.value ?? "continue") === "continue",
      start: Math.max(0, Math.round(Number(this.startInput?.value) || 1)),
    } satisfies PageNumberFormatValues);
    this.hide();
  }

  // The start input answers only when "start at" is the picked mode (Word
  // disables it under continue-from-previous).
  syncMode(): void {
    if (this.startInput)
      this.startInput.disabled = (this.modeGroup?.value ?? "continue") === "continue";
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("pageNumFmt.title", this);
    if (this.formatLabel) this.formatLabel.textContent = t("pageNumFmt.format", this);
    if (this.numberingLabel) this.numberingLabel.textContent = t("pageNumFmt.numbering", this);
    if (this.continueLabel) this.continueLabel.textContent = t("pageNumFmt.continue", this);
    if (this.startLabel) this.startLabel.textContent = t("pageNumFmt.startAt", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenPageNumberFormatDialog;
