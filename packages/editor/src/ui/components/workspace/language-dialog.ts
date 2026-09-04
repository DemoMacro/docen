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
    gap: 14px;
    font-size: 13px;
  }
  fluent-dropdown {
    width: 100%;
    min-width: 0;
  }
  fluent-field {
    align-self: flex-start;
  }
`;

const optionTemplate = html<ProofingLanguage, DocenLanguageDialog>`
  <fluent-option value="${(l) => l.tag}">${(l) => l.name}</fluent-option>
`;

const template = html<DocenLanguageDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="lang-body">
      <fluent-dropdown type="combobox" appearance="outline" ${ref("langDropdown")}>
        <fluent-listbox popover="manual" tabindex="-1">
          ${repeat(() => PROOFING_LANGUAGES, optionTemplate)}
        </fluent-listbox>
        <input
          slot="control"
          role="combobox"
          aria-haspopup="listbox"
          type="combobox"
          size="1"
          style="width:100%;box-sizing:border-box"
          ${ref("langInput")}
        />
      </fluent-dropdown>
      <fluent-field label-position="after">
        <fluent-checkbox slot="input" ${ref("noProofBox")}></fluent-checkbox>
        <label slot="label">${(x) => t("languageDialog.noProof", x)}</label>
      </fluent-field>
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
  // The combobox (value = the picked w:lang tag) and its control input; the
  // Fluent checkbox (checked = w:noProof) — both replace the former raw
  // radio-list + native checkbox.
  @observable langDropdown?: HTMLElement & { value: string | null };
  @observable langInput?: HTMLInputElement;
  @observable noProofBox?: HTMLElement & { checked: boolean };
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
    if (this.langDropdown) this.langDropdown.value = target;
    if (this.noProofBox) this.noProofBox.checked = noProof;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). The combobox
   *  allows free typing, so an unmatched value falls back to the preselected
   *  language rather than stamping an unknown tag. */
  applyLanguage(): void {
    const raw = this.langDropdown?.value ?? null;
    const picked = PROOFING_LANGUAGES.find((l) => l.tag === raw) ?? PROOFING_LANGUAGES[0];
    this.$emit("language:ok", {
      value: picked.tag,
      noProof: this.noProofBox?.checked === true,
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("languageDialog.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenLanguageDialog;
