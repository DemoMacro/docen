import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string };

/** The dialog's values — the visible text and the link address. */
export interface LinkValues {
  text: string;
  href: string;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(400px, 92vw);
  }
  .link-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .field > label {
    white-space: nowrap;
  }
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
`;

const template = html<DocenLinkDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="link-body">
      <div class="field">
        <label ${ref("textLabel")}></label>
        <fluent-text-input ${ref("textInput")}></fluent-text-input>
      </div>
      <div class="field">
        <label ${ref("hrefLabel")}></label>
        <fluent-text-input ${ref("hrefInput")}></fluent-text-input>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyLink()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-link-dialog>` — the hyperlink fields (display text plus address,
 * Word's Insert Link box). Opened by the ribbon Link entry and Ctrl+K; the
 * host prefills the selection's current link via `show(values)` and commits
 * via `link:ok` (an empty address removes the link, Word's Remove
 * Hyperlink). Rides on `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-link-dialog", template, styles })
class DocenLinkDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable textLabel?: HTMLElement;
  @observable textInput?: FluentTextInput;
  @observable hrefLabel?: HTMLElement;
  @observable hrefInput?: FluentTextInput;
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

  /** Prefill from the selection's link (an edit) or its text (an insert). */
  show(values: Partial<LinkValues> = {}): void {
    if (this.textInput) this.textInput.value = values.text ?? "";
    if (this.hrefInput) this.hrefInput.value = values.href ?? "https://";
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). The host
   *  decides what an empty address means. */
  applyLink(): void {
    const href = (this.hrefInput?.value ?? "").trim();
    const text = (this.textInput?.value ?? "").trim();
    this.$emit("link:ok", { text, href });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("linkDialog.title", this);
    if (this.textLabel) this.textLabel.textContent = t("linkDialog.text", this);
    if (this.hrefLabel) this.hrefLabel.textContent = t("linkDialog.address", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenLinkDialog;
