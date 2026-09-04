import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

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
    min-width: 76px;
  }
  .row select,
  .row fluent-text-input {
    flex: 1;
    min-width: 0;
  }
`;

/** The caption's live preview line — Word's dialog keeps the caption shape in
 *  view while the fields change. `seq` stands in for the next SEQ number. */
const template = html<DocenCaptionDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="preview" ${ref("previewEl")}></div>
      <div class="row">
        <label ${ref("labelLabel")}></label>
        <select ${ref("labelSel")} @change="${(x) => x.syncPreview()}"></select>
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
        <select ${ref("positionSel")} @change="${(x) => x.syncPreview()}">
          <option value="below"></option>
          <option value="above"></option>
        </select>
      </div>
      <label class="row">
        <fluent-checkbox ${ref("excludeChk")} @change="${(x) => x.syncPreview()}"></fluent-checkbox>
        <span ${ref("excludeLabel")}></span>
      </label>
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

/** A checkbox widget plus its checked state accessor (fluent-checkbox exposes
 *  `currentChecked`, not the native `checked`). */
type FluentCheckbox = HTMLElement & { currentChecked?: boolean };
type FluentTextInput = HTMLElement & { value: string };

/**
 * `<docen-caption-dialog>` — Word's Insert Caption dialog (题注): a live
 * preview of the caption shape, the label (Figure/Table/Equation — the label
 * written into the document follows the UI language, like Word's), the caption
 * text, the position relative to the anchored item, and the exclude-label
 * flag. Opened with `show()`; commits via `caption:ok`
 * `{ label, text, position, excludeLabel }` or cancels.
 */
@customElement({ name: "docen-caption-dialog", template, styles })
class DocenCaptionDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable previewEl?: HTMLElement;
  @observable labelLabel?: HTMLElement;
  @observable labelSel?: HTMLSelectElement;
  @observable textLabel?: HTMLElement;
  @observable textInput?: FluentTextInput;
  @observable positionLabel?: HTMLElement;
  @observable positionSel?: HTMLSelectElement;
  @observable excludeChk?: FluentCheckbox;
  @observable excludeLabel?: HTMLElement;
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
    if (this.labelSel && this.labelSel.options.length === 0) this.#applyLabels();
    if (this.textInput) this.textInput.value = "";
    if (this.positionSel) this.positionSel.value = "below";
    if (this.excludeChk) this.excludeChk.currentChecked = false;
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
    const label = this.labelSel?.value ?? "";
    const text = this.textInput?.value ?? "";
    const excluded = this.excludeChk?.currentChecked ?? false;
    const head = `${excluded ? "" : `${label} `}1`;
    this.previewEl.textContent = text ? `${head}: ${text}` : head;
  }

  applyCaption(): void {
    const label = this.labelSel?.value ?? "";
    const text = this.textInput?.value ?? "";
    if (!label) return;
    this.$emit("caption:ok", {
      label,
      text,
      position: this.positionSel?.value === "above" ? "above" : "below",
      excludeLabel: this.excludeChk?.currentChecked ?? false,
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("caption.title", this);
    if (this.labelLabel) this.labelLabel.textContent = t("caption.label", this);
    if (this.textLabel) this.textLabel.textContent = t("caption.text", this);
    if (this.positionLabel) this.positionLabel.textContent = t("caption.position", this);
    if (this.excludeLabel) this.excludeLabel.textContent = t("caption.exclude", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.positionSel) {
      const [below, above] = this.positionSel.options;
      if (below) below.textContent = t("caption.below", this);
      if (above) above.textContent = t("caption.above", this);
    }
    // The label options are rebuilt on language change: the option VALUE is
    // the label word written into the document, and Word writes it in the UI
    // language (Chinese Word captions read 图 1, English ones Figure 1).
    if (this.labelSel)
      this.labelSel.replaceChildren(
        new Option(t("caption.figure", this), t("caption.figure", this)),
        new Option(t("caption.table", this), t("caption.table", this)),
        new Option(t("caption.equation", this), t("caption.equation", this)),
      );
  }
}

export default DocenCaptionDialog;
