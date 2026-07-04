import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const LANGS = [
  { id: "zh-CN", key: "options.lang.zh" },
  { id: "en", key: "options.lang.en" },
] as const;

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
 * Note: `fluent-radio` renders no text of its own (see registry.ts), so each
 * radio is wrapped in a `fluent-field` with a slotted `<label>` — the same
 * pattern the properties panel uses.
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
  #langLocal = "zh-CN";

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
    this.#langLocal = this.locale ?? "zh-CN";
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
    if (this.dialogEl) this.dialogEl.heading = t("options.title");
    if (this.okBtn) this.okBtn.textContent = t("options.ok");
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel");
  }

  #renderLanguage(): void {
    if (!this.contentEl) return;
    this.contentEl.replaceChildren();
    const heading = document.createElement("div");
    heading.className = "opt-heading";
    heading.textContent = t("options.language");
    this.contentEl.append(heading);
    const group = document.createElement("fluent-radio-group");
    group.setAttribute("name", "opt-lang");
    group.setAttribute("orientation", "vertical");
    for (const l of LANGS) {
      const id = `opt-lang-${l.id}`;
      const fieldEl = document.createElement("fluent-field");
      fieldEl.setAttribute("label-position", "after");
      const label = document.createElement("label");
      label.slot = "label";
      label.htmlFor = id;
      label.id = `${id}--label`;
      label.textContent = t(l.key);
      const radio = document.createElement("fluent-radio") as HTMLElement & {
        value: string;
      };
      radio.slot = "input";
      radio.id = id;
      radio.setAttribute("name", "opt-lang");
      radio.value = l.id;
      radio.setAttribute("aria-labelledby", `${id}--label`);
      if (l.id === this.#langLocal) radio.setAttribute("checked", "");
      radio.addEventListener("change", () => {
        this.#langLocal = l.id;
      });
      fieldEl.append(label, radio);
      group.append(fieldEl);
    }
    this.contentEl.append(group);
  }
}

export default DocenOptionsDialog;
