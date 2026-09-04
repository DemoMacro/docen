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

/** One proofing language — the tag written into w:lang and the name the
 *  dialog lists (the language's own name, like Word's list). */
export interface ProofingLanguage {
  tag: string;
  name: string;
}

/** The built-in proofing languages (Word's list, most common first). */
export const PROOFING_LANGUAGES: ProofingLanguage[] = [
  { tag: "zh-CN", name: "中文（中国）" },
  { tag: "zh-TW", name: "中文（台灣）" },
  { tag: "en-US", name: "English (US)" },
  { tag: "en-GB", name: "English (UK)" },
  { tag: "ja-JP", name: "日本語 (日本)" },
  { tag: "ko-KR", name: "한국어(대한민국)" },
  { tag: "fr-FR", name: "Français (France)" },
  { tag: "de-DE", name: "Deutsch (Deutschland)" },
  { tag: "es-ES", name: "Español (España)" },
  { tag: "ru-RU", name: "Русский (Россия)" },
];

/** The display name of a w:lang tag (falls back to the tag itself). */
export function proofingLanguageName(tag: string): string {
  return PROOFING_LANGUAGES.find((l) => l.tag === tag)?.name ?? tag;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(320px, 92vw);
  }
  .lang-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .list {
    max-height: 260px;
    overflow: auto;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .choice {
    display: flex;
    align-items: center;
    gap: 8px;
    cursor: pointer;
    padding: 1px 2px;
  }
  .no-proof {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenLanguageDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="lang-body">
      <div class="list">
        ${repeat(
          () => PROOFING_LANGUAGES,
          html`<label class="choice">
            <input type="radio" name="proofing-lang" value="${(l) => l.tag}" />
            <span>${(l) => l.name}</span>
          </label>`,
        )}
      </div>
      <label class="no-proof">
        <input type="checkbox" ${ref("noProofBox")} />
        <span>${(x) => t("languageDialog.noProof", x)}</span>
      </label>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyLanguage()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-language-dialog>` — Word's Language dialog (Review → Language →
 * Set Proofing Language, and the status-bar language item): pick a proofing
 * language for the selection (written to the runs' w:lang) and optionally
 * mark it "do not check spelling" (w:noProof). Opened with `show(tag)` —
 * the selection's current language preselected; commits via `language:ok`
 * `{ value, noProof }`.
 */
@customElement({ name: "docen-language-dialog", template, styles })
class DocenLanguageDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable noProofBox?: HTMLInputElement;
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

  /** Open with `tag` preselected (the first entry when null/unlisted) and the
   *  no-proof box mirroring the selection's current state. */
  show(tag: string | null, noProof = false): void {
    const target = PROOFING_LANGUAGES.some((l) => l.tag === tag) ? tag : PROOFING_LANGUAGES[0].tag;
    for (const radio of this.shadowRoot?.querySelectorAll<HTMLInputElement>(
      'input[name="proofing-lang"]',
    ) ?? []) {
      radio.checked = radio.value === target;
    }
    if (this.noProofBox) this.noProofBox.checked = noProof;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyLanguage(): void {
    const picked = this.shadowRoot?.querySelector<HTMLInputElement>(
      'input[name="proofing-lang"]:checked',
    );
    if (!picked) return;
    this.$emit("language:ok", { value: picked.value, noProof: this.noProofBox?.checked === true });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("languageDialog.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenLanguageDialog;
