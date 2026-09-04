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

/** The formats the dialog offers — the clipboard lanes this editor can read
 *  (Word's dialog lists more: RTF, pictures; the browser clipboard only
 *  reliably carries styled HTML and plain text). */
export type PasteSpecialFormat = "html" | "text";

const FORMATS: Array<{ value: PasteSpecialFormat; key: string }> = [
  { value: "html", key: "pasteSpecial.html" },
  { value: "text", key: "pasteSpecial.text" },
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(300px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  fluent-radio-group {
    align-items: flex-start;
    gap: 6px;
  }
  .as {
    color: var(--docen-color-text-2, #616161);
  }
`;

const formatTemplate = html<(typeof FORMATS)[number], DocenPasteSpecialDialog>`
  <fluent-field label-position="after">
    <fluent-radio slot="input" value="${(f) => f.value}"></fluent-radio>
    <label slot="label"><span class="format-label" data-key="${(f) => f.key}"></span></label>
  </fluent-field>
`;

const template = html<DocenPasteSpecialDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <span class="as">${(x) => t("pasteSpecial.as", x)}</span>
      <fluent-radio-group orientation="vertical" ${ref("formatGroup")}>
        ${repeat(() => FORMATS, formatTemplate)}
      </fluent-radio-group>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyFormat()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-paste-special-dialog>` — Word's Paste Special dialog (Home →
 * Paste ▾ → Paste Special…): picks the format the clipboard content pastes
 * as. Commits via `paste-special:ok` with a {@link PasteSpecialFormat};
 * the actual paste (and its clipboard read) stays with the host.
 */
@customElement({ name: "docen-paste-special-dialog", template, styles })
class DocenPasteSpecialDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  // The group's `value` mirrors the checked radio's `value` ("html" | "text").
  @observable formatGroup?: HTMLElement & { value: string | null };
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

  /** Open with formatted HTML pre-picked (Word's default). */
  show(): void {
    if (this.formatGroup) this.formatGroup.value = FORMATS[0].value;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyFormat(): void {
    const picked = this.formatGroup?.value ?? FORMATS[0].value;
    this.$emit("paste-special:ok", picked as PasteSpecialFormat);
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("pasteSpecial.title", this);
    for (const el of this.shadowRoot?.querySelectorAll<HTMLElement>(".format-label") ?? [])
      el.textContent = t(el.dataset.key ?? "", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenPasteSpecialDialog;
