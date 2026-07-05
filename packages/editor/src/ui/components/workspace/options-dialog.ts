import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { availableLanguages, observeLang, resolveLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(380px, 92vw);
  }
  .opt-body {
    padding: 8px 4px 4px;
  }
  .opt-heading {
    font-weight: 600;
    margin-block-end: 8px;
  }
  /* Native <select> — dependency-free and accessible; Office's language picker
     is a plain dropdown too. Styled against the docen token palette so it
     tracks light/dark along with the rest of the shell. */
  .opt-lang-select {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    background: var(--docen-color-background-1, #fff);
    color: var(--docen-color-text-1, #242424);
    font-size: 13px;
    font-family: inherit;
    cursor: pointer;
  }
`;

const template = html<DocenOptionsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="opt-body" ${ref("contentEl")}></div>
    <div slot="action" class="opt-actions">
      <fluent-button appearance="stealth" ${ref("cancelBtn")}></fluent-button>
      <fluent-button appearance="accent" ${ref("okBtn")}></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-options-dialog locale="…">` — MS Office "选项" dialog. v1 keeps it
 * minimal: just the UI-language picker (the rest of the Office categories land
 * later). Rides on `<docen-dialog>` for the modal shell (backdrop / Esc / show).
 *
 * The host seeds the current language via `locale`, calls `show()`, and listens
 * for `options:ok { lang }` (确定). Cancel / Esc just close. State commits
 * atomically on OK (Office behavior — not live).
 *
 * The language list is data-driven: {@link availableLanguages} reads every
 * registered tag, so a new locale added via `registerTranslation` (or an
 * add-in's `localizationInfo`) appears here with no further wiring.
 */
@customElement({ name: "docen-options-dialog", template, styles })
class DocenOptionsDialog extends FASTElement {
  // `locale`, not `lang` — HTMLElement already declares `lang`, so @attr lang
  // clashes with the base property (TS2416).
  @attr locale?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable contentEl?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  #unobserveLang?: () => void;
  #langLocal = "";

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderLanguage();
    this.#applyLabels();
    this.okBtn?.addEventListener("click", this.#onOk);
    this.cancelBtn?.addEventListener("click", this.#onCancel);
    this.#unobserveLang = observeLang(() => {
      this.#renderLanguage();
      this.#applyLabels();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.okBtn?.removeEventListener("click", this.#onOk);
    this.cancelBtn?.removeEventListener("click", this.#onCancel);
    super.disconnectedCallback();
  }

  show(): void {
    this.#langLocal = this.locale ?? resolveLang(this);
    this.#renderLanguage();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  readonly #onOk = (): void => {
    this.dispatchEvent(
      new CustomEvent("options:ok", {
        bubbles: true,
        composed: true,
        detail: { lang: this.#langLocal },
      }),
    );
    this.hide();
  };

  readonly #onCancel = (): void => {
    this.hide();
  };

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("options.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }

  #renderLanguage(): void {
    if (!this.contentEl) return;
    this.contentEl.replaceChildren();
    const heading = document.createElement("div");
    heading.className = "opt-heading";
    heading.textContent = t("options.language", this);
    this.contentEl.append(heading);
    // Pick the selected option from the live `#langLocal` (set by show()) or
    // the host's current locale — so the dialog opens pointing at the right row
    // even before the user touches it.
    const current = this.#langLocal || this.locale || resolveLang(this);
    const select = document.createElement("select");
    select.className = "opt-lang-select";
    select.setAttribute("aria-label", t("options.language", this));
    for (const l of availableLanguages()) {
      const option = document.createElement("option");
      option.value = l.languageTag;
      option.textContent = l.$name ?? l.languageTag;
      if (l.languageTag === current) option.selected = true;
      select.append(option);
    }
    select.addEventListener("change", () => {
      this.#langLocal = select.value;
    });
    this.contentEl.append(select);
  }
}

export default DocenOptionsDialog;
