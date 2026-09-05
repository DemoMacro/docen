import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { MergeRecipients } from "../../../document/commands/mail-merge";
import { parseRecipients } from "../../../document/commands/mail-merge";
import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(520px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    font-size: 13px;
  }
  .hint {
    color: var(--colorNeutralForeground2, #616161);
  }
  fluent-textarea {
    width: 100%;
    min-height: 160px;
    font-family: inherit;
  }
  .preview {
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    padding: 6px 8px;
    color: var(--colorNeutralForeground2, #616161);
  }
`;

const template = html<DocenRecipientsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <span class="hint">${(x) => t("recipients.hint", x)}</span>
      <fluent-textarea ${ref("dataInput")} spellcheck="false"></fluent-textarea>
      <div class="preview" ${ref("previewEl")}></div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.commit()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

type FluentTextInput = HTMLElement & { value: string };

/**
 * `<docen-recipients-dialog>` — Word's Select Recipients over a pasted data
 *  source: a CSV/TSV grid (header row + records) typed or pasted in, parsed
 *  live so the count preview confirms before commit. Commits via
 *  `recipients:ok` `{ recipients }` — null when the text parses to nothing
 *  (which clears the document's data source).
 */
@customElement({ name: "docen-recipients-dialog", template, styles })
class DocenRecipientsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable dataInput?: FluentTextInput;
  @observable previewEl?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
    this.dataInput?.addEventListener("input", () => this.#updatePreview());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Open prefilled with the document's current data source (edit mode). */
  show(recipients: MergeRecipients | null): void {
    if (this.dataInput) {
      this.dataInput.value = recipients
        ? [recipients.headers.join("\t"), ...recipients.rows.map((r) => r.join("\t"))].join("\n")
        : "";
    }
    this.#updatePreview();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  commit(): void {
    this.$emit("recipients:ok", { recipients: parseRecipients(this.dataInput?.value ?? "") });
    this.hide();
  }

  #updatePreview(): void {
    if (!this.previewEl) return;
    const parsed = parseRecipients(this.dataInput?.value ?? "");
    this.previewEl.textContent = parsed
      ? t("recipients.preview", this)
          .replace("{0}", String(parsed.rows.length))
          .replace("{1}", parsed.headers.join(", "))
      : t("recipients.previewEmpty", this);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("recipients.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    this.#updatePreview();
  }
}

export default DocenRecipientsDialog;
