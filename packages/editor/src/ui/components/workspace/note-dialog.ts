import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .body fluent-textarea {
    width: 100%;
    box-sizing: border-box;
  }
`;

const template = html<DocenNoteDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <label ${ref("textLabel")}></label>
      <fluent-textarea ${ref("textArea")} block rows="5" spellcheck="false"></fluent-textarea>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyNote()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

type FluentTextarea = HTMLTextAreaElement & HTMLElement;

/**
 * `<docen-note-dialog>` — the footnote/endnote text editor (Word edits the
 * note in place at the page bottom; the canvas note area is projected, so the
 * editor surface is this dialog). `show(kind, text)` prefills the note's text
 * — one textarea line per note paragraph; commits via `note:ok` `{ text }`.
 */
@customElement({ name: "docen-note-dialog", template, styles })
class DocenNoteDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable textLabel?: HTMLElement;
  @observable textArea?: FluentTextarea;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;
  #kind: "footnote" | "endnote" = "footnote";

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

  show(kind: "footnote" | "endnote", text: string): void {
    this.#kind = kind;
    this.#applyLabels();
    if (this.textArea) this.textArea.value = text;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  applyNote(): void {
    this.$emit("note:ok", { kind: this.#kind, text: this.textArea?.value ?? "" });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl)
      this.dialogEl.heading =
        this.#kind === "endnote" ? t("note.endnote-title", this) : t("note.footnote-title", this);
    if (this.textLabel) this.textLabel.textContent = t("note.text", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenNoteDialog;
