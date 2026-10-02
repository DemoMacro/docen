import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { formatSeqNumber } from "../../../document/fields";
import { observeLang, t } from "../../i18n/localize";
import { listboxOf, opt, pick, pickedValue, type FluentDropdown } from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(400px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .preview {
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    background: var(--colorNeutralBackground1, #fff);
    padding: 6px 10px;
    min-height: 20px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .row > label {
    min-width: 96px;
  }
  .row fluent-dropdown,
  .row fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .row fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
`;

/** The caption's live preview line — Word's dialog keeps the caption shape in
 *  view while the fields change. The number stands in for the next SEQ value
 *  (the chapter segment stands in for the current chapter). */
const template = html<DocenCaptionDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="preview" ${ref("previewEl")}></div>
      <div class="row">
        <label ${ref("labelLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("labelSel")}
          @change="${(x) => x.syncPreview()}"
        >
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
      </div>
      <div class="row">
        <label ${ref("textLabel")}></label>
        <fluent-text-input
          ${ref("textInput")}
          @input="${(x) => x.syncPreview()}"
          spellcheck="false"
        ></fluent-text-input>
      </div>
      <div class="row">
        <label ${ref("positionLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("positionSel")}
          @change="${(x) => x.syncPreview()}"
        >
          <fluent-listbox popover="manual" tabindex="-1">
            <fluent-option value="below"></fluent-option>
            <fluent-option value="above"></fluent-option>
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
      </div>
      <label class="row">
        <fluent-checkbox ${ref("excludeChk")} @change="${(x) => x.syncPreview()}"></fluent-checkbox>
        <span ${ref("excludeLabel")}></span>
      </label>
      <div class="row">
        <label ${ref("formatLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("formatSel")}
          @change="${(x) => x.syncPreview()}"
        >
          <fluent-listbox popover="manual" tabindex="-1">
            <fluent-option value="ARABIC">1, 2, 3, …</fluent-option>
            <fluent-option value="ROMAN">I, II, III, …</fluent-option>
            <fluent-option value="roman">i, ii, iii, …</fluent-option>
            <fluent-option value="ALPHABETIC">A, B, C, …</fluent-option>
            <fluent-option value="alphabetic">a, b, c, …</fluent-option>
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
      </div>
      <label class="row">
        <fluent-checkbox ${ref("chapterChk")} @change="${(x) => x.syncPreview()}"></fluent-checkbox>
        <span ${ref("chapterLabel")}></span>
      </label>
      <div class="row">
        <label ${ref("chapterStyleLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("chapterStyleSel")}
          @change="${(x) => x.syncPreview()}"
        >
          <fluent-listbox popover="manual" tabindex="-1">
            <fluent-option value="1"></fluent-option>
            <fluent-option value="2"></fluent-option>
            <fluent-option value="3"></fluent-option>
            <fluent-option value="4"></fluent-option>
            <fluent-option value="5"></fluent-option>
            <fluent-option value="6"></fluent-option>
            <fluent-option value="7"></fluent-option>
            <fluent-option value="8"></fluent-option>
            <fluent-option value="9"></fluent-option>
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
      </div>
      <div class="row">
        <label ${ref("separatorLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("separatorSel")}
          @change="${(x) => x.syncPreview()}"
        >
          <fluent-listbox popover="manual" tabindex="-1">
            <fluent-option value="hyphen"></fluent-option>
            <fluent-option value="period"></fluent-option>
            <fluent-option value="colon"></fluent-option>
            <fluent-option value="emDash"></fluent-option>
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
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyCaption()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A checkbox widget. The rendered state follows `checked`; `currentChecked`
 *  is a separate slot only user clicks keep in sync. */
type FluentCheckbox = HTMLElement & { checked?: boolean };
type FluentTextInput = HTMLElement & { value: string };

/** The separator ladder in template order — the tokens mirror w:caption@w:sep
 *  (Word's "Use separator" list), the preview needs their character. */
const SEPARATORS: readonly { token: string; char: string; key: string }[] = [
  { token: "hyphen", char: "-", key: "caption.sep.hyphen" },
  { token: "period", char: ".", key: "caption.sep.period" },
  { token: "colon", char: ":", key: "caption.sep.colon" },
  { token: "emDash", char: "\u2014", key: "caption.sep.emDash" },
];

/**
 * `<docen-caption-dialog>` — Word's Insert Caption dialog (题注): a live
 * preview of the caption shape, the label (Figure/Table/Equation — the label
 * written into the document follows the UI language, like Word's), the caption
 * text, the position relative to the anchored item, the exclude-label flag,
 * the number format, and the chapter-number shape ("Include chapter number" +
 * chapter-start heading level + separator). Opened with `show()`; commits via
 * `caption:ok` `{ label, text, position, excludeLabel, format, chapterNumber,
 * heading, sep }` or cancels.
 */
@customElement({ name: "docen-caption-dialog", template, styles })
class DocenCaptionDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable previewEl?: HTMLElement;
  @observable labelLabel?: HTMLElement;
  @observable labelSel?: FluentDropdown;
  @observable textLabel?: HTMLElement;
  @observable textInput?: FluentTextInput;
  @observable positionLabel?: HTMLElement;
  @observable positionSel?: FluentDropdown;
  @observable excludeChk?: FluentCheckbox;
  @observable excludeLabel?: HTMLElement;
  @observable formatLabel?: HTMLElement;
  @observable formatSel?: FluentDropdown;
  @observable chapterChk?: FluentCheckbox;
  @observable chapterLabel?: HTMLElement;
  @observable chapterStyleLabel?: HTMLElement;
  @observable chapterStyleSel?: FluentDropdown;
  @observable separatorLabel?: HTMLElement;
  @observable separatorSel?: FluentDropdown;
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

  show(): void {
    this.#applyLabels();
    if (this.textInput) this.textInput.value = "";
    pick(this.positionSel, "below");
    pick(this.formatSel, "ARABIC");
    pick(this.chapterStyleSel, "1");
    pick(this.separatorSel, "hyphen");
    if (this.excludeChk) this.excludeChk.checked = false;
    if (this.chapterChk) this.chapterChk.checked = false;
    this.syncPreview();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** The caption shape as it will land in the document (the host knows the
   *  real next SEQ number; the preview always shows 1). */
  syncPreview(): void {
    if (!this.previewEl) return;
    const label = pickedValue(this.labelSel) ?? "";
    const text = this.textInput?.value ?? "";
    const excluded = this.excludeChk?.checked ?? false;
    const format = this.formatSel?.value || "ARABIC";
    const number = formatSeqNumber(format, 1);
    // A chapter-numbered caption reads "chapter + separator + number" (the
    // chapter stands in at 1; the separator is the picked one).
    const separator = SEPARATORS.find((s) => s.token === this.separatorSel?.value)?.char ?? "-";
    const head = `${excluded ? "" : `${label} `}${
      this.chapterChk?.checked ? `1${separator}` : ""
    }${number}`;
    this.previewEl.textContent = text ? `${head}: ${text}` : head;
  }

  applyCaption(): void {
    const label = pickedValue(this.labelSel) ?? "";
    const text = this.textInput?.value ?? "";
    if (!label) return;
    this.$emit("caption:ok", {
      label,
      text,
      position: pickedValue(this.positionSel) === "above" ? "above" : "below",
      excludeLabel: this.excludeChk?.checked ?? false,
      format: this.formatSel?.value || "ARABIC",
      chapterNumber: this.chapterChk?.checked ?? false,
      heading: Number(this.chapterStyleSel?.value) || 1,
      sep: this.separatorSel?.value || "hyphen",
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("caption.title", this);
    if (this.labelLabel) this.labelLabel.textContent = t("caption.label", this);
    if (this.textLabel) this.textLabel.textContent = t("caption.text", this);
    if (this.positionLabel) this.positionLabel.textContent = t("caption.position", this);
    if (this.excludeLabel) this.excludeLabel.textContent = t("caption.exclude", this);
    if (this.formatLabel) this.formatLabel.textContent = t("caption.format", this);
    if (this.chapterLabel) this.chapterLabel.textContent = t("caption.chapter", this);
    if (this.chapterStyleLabel)
      this.chapterStyleLabel.textContent = t("caption.chapterStyle", this);
    if (this.separatorLabel) this.separatorLabel.textContent = t("caption.separator", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.positionSel) {
      const [below, above] = this.positionSel.querySelectorAll("fluent-option");
      if (below) below.textContent = t("caption.below", this);
      if (above) above.textContent = t("caption.above", this);
    }
    // The chapter-start choices are the heading styles, localized like Word's
    // ("Heading 1" / "标题 1"); values stay the 1-9 levels.
    if (this.chapterStyleSel) {
      const options = this.chapterStyleSel.querySelectorAll("fluent-option");
      options.forEach((option, index) => {
        option.textContent = t(`caption.heading${index + 1}`, this);
      });
    }
    if (this.separatorSel) {
      const options = this.separatorSel.querySelectorAll("fluent-option");
      options.forEach((option, index) => {
        const key = SEPARATORS[index]?.key;
        if (key) option.textContent = t(key, this);
      });
    }
    // The label options are rebuilt on language change: the option VALUE is
    // the label word written into the document, and Word writes it in the UI
    // language (Chinese Word captions read 图 1, English ones Figure 1). The
    // first label is the pick, like a native select's initial selection.
    if (this.labelSel) {
      const figure = t("caption.figure", this);
      listboxOf(this.labelSel)?.replaceChildren(
        opt(figure, figure),
        opt(t("caption.table", this), t("caption.table", this)),
        opt(t("caption.equation", this), t("caption.equation", this)),
      );
      pick(this.labelSel, figure);
    }
  }
}

export default DocenCaptionDialog;
