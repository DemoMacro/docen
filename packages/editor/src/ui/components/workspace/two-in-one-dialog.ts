import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(320px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .field {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .field input[type="text"] {
    width: 100%;
    box-sizing: border-box;
  }
  .preview {
    border: 1px dashed var(--neutral-stroke-rest, #d1d1d1);
    border-radius: 3px;
    padding: 8px 0;
    text-align: center;
    font-size: 15px;
    min-height: 20px;
    user-select: none;
  }
  .check {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenTwoInOneDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <label class="field">
        <span>${(x) => t("twoInOne.text", x)}</span>
        <input type="text" ${ref("textEl")} spellcheck="false" @input="${(x) => x.syncPreview()}" />
      </label>
      <div class="preview" ${ref("previewEl")}></div>
      <label class="check">
        <input type="checkbox" ${ref("bracketsEl")} @change="${(x) => x.syncPreview()}" />
        <span>${(x) => t("twoInOne.brackets", x)}</span>
      </label>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.apply()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-two-in-one-dialog>` — Word's Two Lines in One dialog (双行合一,
 * Home → Paragraph → Chinese Layout): the text to pack into two lines (a
 * space marks the split, Word's rule) and the bracket pair option. Opened
 * with `show(text?)` — the selection's text prefilled; commits via
 * `two-in-one:ok` `{ text, brackets }` or cancels.
 */
@customElement({ name: "docen-two-in-one-dialog", template, styles })
class DocenTwoInOneDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable textEl?: HTMLInputElement;
  @observable previewEl?: HTMLDivElement;
  @observable bracketsEl?: HTMLInputElement;
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

  /** Open with the selection's text (and bracket state) prefilled. */
  show(text = "", brackets = false): void {
    if (this.textEl) this.textEl.value = text;
    if (this.bracketsEl) this.bracketsEl.checked = brackets;
    this.syncPreview();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  apply(): void {
    const text = this.textEl?.value ?? "";
    if (!text) return;
    this.$emit("two-in-one:ok", { text, brackets: this.bracketsEl?.checked === true });
    this.hide();
  }

  syncPreview(): void {
    if (!this.previewEl || !this.textEl) return;
    // Word's dialog preview: the split at the first space (else even), the
    // projection's own rule.
    const text = this.textEl.value;
    const at = text.indexOf(" ");
    const first = at >= 0 ? text.slice(0, at) : text.slice(0, Math.ceil(text.length / 2));
    const second = at >= 0 ? text.slice(at + 1) : text.slice(Math.ceil(text.length / 2));
    this.previewEl.replaceChildren(
      document.createTextNode(
        `${this.bracketsEl?.checked ? "(" : ""}${first}${this.bracketsEl?.checked ? ")" : ""}`,
      ),
      document.createElement("br"),
      document.createTextNode(
        `${this.bracketsEl?.checked ? "(" : ""}${second}${this.bracketsEl?.checked ? ")" : ""}`,
      ),
    );
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("twoInOne.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenTwoInOneDialog;
