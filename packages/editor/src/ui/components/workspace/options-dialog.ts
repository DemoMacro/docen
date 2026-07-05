import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { availableLanguages, observeLang, resolveLang, t } from "../../i18n/localize";
import type { LanguageOption } from "../../i18n/localize";

// Per-instance CSS anchor name so the dropdown's listbox popover floats under
// its own control — without it, Fluent's default strands the popover at the
// viewport corner (same race as <docen-ribbon-combobox>).
let seq = 0;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(380px, 92vw);
  }
  .opt-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .opt-heading {
    font-weight: 600;
  }
  fluent-dropdown {
    width: 100%;
    min-width: 0;
  }
  input {
    width: 100%;
    box-sizing: border-box;
  }
`;

const optionTemplate = html<LanguageOption, DocenOptionsDialog>`
  <fluent-option value="${(o) => o.languageTag}">${(o) => o.$name ?? o.languageTag}</fluent-option>
`;

const template = html<DocenOptionsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="opt-body">
      <div class="opt-heading" ${ref("headingEl")}></div>
      <fluent-dropdown
        type="combobox"
        appearance="outline"
        part="dropdown"
        ${ref("dropdown")}
        @change="${(x) => x.onLangChange()}"
      >
        <fluent-listbox popover="manual" tabindex="-1" part="listbox" ${ref("listbox")}>
          ${repeat((x) => x.languages, optionTemplate)}
        </fluent-listbox>
        <input
          slot="control"
          role="combobox"
          aria-haspopup="listbox"
          type="combobox"
          part="input"
          size="1"
          style="width:100%;box-sizing:border-box"
          ${ref("input")}
        />
      </fluent-dropdown>
    </div>
    <div slot="action" class="opt-actions">
      <fluent-button
        appearance="stealth"
        ${ref("cancelBtn")}
        @click="${(x) => x.hide()}"
      ></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.onOk()}"
      ></fluent-button>
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
 * The picker is a `<fluent-dropdown type="combobox">` — typeable, so a long
 * locale list can be searched — over {@link availableLanguages}, so a new locale
 * added via `registerTranslation` (or an add-in's `localizationInfo`) appears
 * here with no further wiring.
 */
@customElement({ name: "docen-options-dialog", template, styles })
class DocenOptionsDialog extends FASTElement {
  // `locale`, not `lang` — HTMLElement already declares `lang`, so @attr lang
  // clashes with the base property (TS2416).
  @attr locale?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable headingEl?: HTMLElement;
  @observable dropdown?: HTMLElement;
  @observable listbox?: HTMLElement;
  @observable input?: HTMLInputElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  /** Pickable locales — refreshed when a locale is registered at runtime. */
  @observable languages: readonly LanguageOption[] = availableLanguages();

  readonly popoverId = `opt-lang-${++seq}`;
  readonly popoverAnchor = `--${this.popoverId}`;

  #unobserveLang?: () => void;
  #langLocal = "";

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    // CSS Anchor Positioning: pin the listbox popover to this dropdown.
    if (this.listbox) {
      this.listbox.id = this.popoverId;
      this.listbox.style.positionAnchor = this.popoverAnchor;
    }
    if (this.input) this.input.setAttribute("aria-controls", this.popoverId);
    if (this.dropdown) this.dropdown.style.anchorName = this.popoverAnchor;
    this.#unobserveLang = observeLang(() => {
      this.languages = availableLanguages();
      this.#applyLabels();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(): void {
    this.#langLocal = this.locale ?? resolveLang(this);
    this.#syncSelection();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  readonly onOk = (): void => {
    this.dispatchEvent(
      new CustomEvent("options:ok", {
        bubbles: true,
        composed: true,
        detail: { lang: this.#langLocal },
      }),
    );
    this.hide();
  };

  onLangChange(): void {
    const value = (this.dropdown as unknown as { value: string | null })?.value;
    if (value) this.#langLocal = value;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("options.title", this);
    if (this.headingEl) this.headingEl.textContent = t("options.language", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }

  /** Select the option matching #langLocal once the combobox's control +
   *  listbox are both ready. fluent-dropdown's connectedCallback enqueues
   *  insertControl() (drops the seeded <input>, renders its own bound to an
   *  internal observable) and the listbox slots async, so selectOption() —
   *  which both marks the option selected and writes its text to the control —
   *  must wait. Mirrors <docen-ribbon-combobox>'s syncValue. */
  #syncSelection(): void {
    const dd = this.dropdown as unknown as {
      listbox?: unknown;
      control?: HTMLInputElement;
      selectOption(i: number): void;
    };
    const lb = this.listbox;
    if (!lb) return;
    let tries = 0;
    const apply = (): void => {
      if (!this.isConnected || tries++ > 10) return;
      if (!dd.listbox || !dd.control) {
        requestAnimationFrame(apply);
        return;
      }
      const opts = lb.querySelectorAll("fluent-option");
      let idx = -1;
      opts.forEach((opt, i) => {
        if (opt.getAttribute("value") === this.#langLocal) idx = i;
      });
      if (idx >= 0) dd.selectOption(idx);
    };
    requestAnimationFrame(apply);
  }
}

export default DocenOptionsDialog;
