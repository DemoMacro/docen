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
import { builtinThemes, resolveTheme } from "../../theme";

interface ThemeOption {
  key: string;
  label: string;
}

// Per-instance CSS anchor names so each dropdown's listbox popover floats
// under its own control — without it, Fluent's default strands the popover at
// the viewport corner (same race as <docen-ribbon-combobox>).
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
    gap: 12px;
  }
  .opt-field {
    display: flex;
    flex-direction: column;
    gap: 4px;
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

const themeOptionTemplate = html<ThemeOption, DocenOptionsDialog>`
  <fluent-option value="${(o) => o.key}">${(o) => o.label}</fluent-option>
`;

const template = html<DocenOptionsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="opt-body">
      <div class="opt-field">
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
      <div class="opt-field">
        <div class="opt-heading" ${ref("themeHeadingEl")}></div>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          part="theme-dropdown"
          ${ref("themeDropdown")}
          @change="${(x) => x.onThemeChange()}"
        >
          <fluent-listbox
            popover="manual"
            tabindex="-1"
            part="theme-listbox"
            ${ref("themeListbox")}
          >
            ${repeat((x) => x.themeOptions, themeOptionTemplate)}
          </fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            part="theme-input"
            size="1"
            style="width:100%;box-sizing:border-box"
            ${ref("themeInput")}
          />
        </fluent-dropdown>
      </div>
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

type ComboboxLike = {
  listbox?: unknown;
  control?: HTMLInputElement;
  selectOption(i: number): void;
};

/**
 * `<docen-options-dialog locale="…" theme="…">` — MS Office "Options" dialog.
 * v1 carries the two host-level prefs: UI language (over {@link availableLanguages},
 * so a locale added via `registerTranslation` or an add-in's `localizationInfo`
 * appears here with no further wiring) and theme (the built-in Fluent web /
 * teams / high-contrast themes, plus any registered via registerTheme). Rides
 * on `<docen-dialog>` for
 * the modal shell (backdrop / Esc / show).
 *
 * The host seeds the current values via `locale` / `theme`, calls `show()`, and
 * listens for `options:ok { lang, theme }` (确定). Cancel / Esc just close.
 * State commits atomically on OK (Office behavior — not live).
 *
 * Both pickers are `<fluent-dropdown type="combobox">` — typeable, so a long
 * locale list can be searched.
 */
@customElement({ name: "docen-options-dialog", template, styles })
class DocenOptionsDialog extends FASTElement {
  // `locale`, not `lang` — HTMLElement already declares `lang`, so @attr lang
  // clashes with the base property (TS2416).
  @attr locale?: string;
  @attr theme?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable headingEl?: HTMLElement;
  @observable dropdown?: HTMLElement;
  @observable listbox?: HTMLElement;
  @observable input?: HTMLInputElement;
  @observable themeHeadingEl?: HTMLElement;
  @observable themeDropdown?: HTMLElement;
  @observable themeListbox?: HTMLElement;
  @observable themeInput?: HTMLInputElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  /** Pickable locales — refreshed when a locale is registered at runtime. */
  @observable languages: readonly LanguageOption[] = availableLanguages();
  /** Pickable themes — labels re-resolve when the locale changes. */
  @observable themeOptions: readonly ThemeOption[] = [];

  readonly popoverId = `opt-lang-${++seq}`;
  readonly popoverAnchor = `--${this.popoverId}`;
  readonly themePopoverId = `opt-theme-${++seq}`;
  readonly themePopoverAnchor = `--${this.themePopoverId}`;

  #unobserveLang?: () => void;
  #langLocal = "";
  #themeLocal = "";

  connectedCallback(): void {
    super.connectedCallback();
    this.themeOptions = this.#computeThemeOptions();
    this.#applyLabels();
    this.#wireCombobox(this.dropdown, this.listbox, this.input, this.popoverId, this.popoverAnchor);
    this.#wireCombobox(
      this.themeDropdown,
      this.themeListbox,
      this.themeInput,
      this.themePopoverId,
      this.themePopoverAnchor,
    );
    this.#unobserveLang = observeLang(() => {
      this.languages = availableLanguages();
      this.themeOptions = this.#computeThemeOptions();
      this.#applyLabels();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(): void {
    // Refresh in case a theme was registered since the dialog last opened.
    this.themeOptions = this.#computeThemeOptions();
    this.#langLocal = this.locale ?? resolveLang(this);
    this.#themeLocal = resolveTheme(this.theme);
    this.#syncCombobox(
      this.dropdown as unknown as ComboboxLike | undefined,
      this.listbox,
      this.#langLocal,
    );
    this.#syncCombobox(
      this.themeDropdown as unknown as ComboboxLike | undefined,
      this.themeListbox,
      this.#themeLocal,
    );
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
        detail: { lang: this.#langLocal, theme: this.#themeLocal },
      }),
    );
    this.hide();
  };

  onLangChange(): void {
    const value = (this.dropdown as unknown as { value: string | null })?.value;
    if (value) this.#langLocal = value;
  }

  onThemeChange(): void {
    const value = (this.themeDropdown as unknown as { value: string | null })?.value;
    if (value) this.#themeLocal = value;
  }

  #computeThemeOptions(): ThemeOption[] {
    return [...builtinThemes.keys()].map((key) => ({ key, label: t(`theme.${key}`, this) }));
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("options.title", this);
    if (this.headingEl) this.headingEl.textContent = t("options.language", this);
    if (this.themeHeadingEl) this.themeHeadingEl.textContent = t("options.theme", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }

  /** Pin a listbox popover to its dropdown via CSS Anchor Positioning and wire
   *  the input's aria-controls. Each dropdown needs its own anchor name or the
   *  popovers collide at the viewport corner. */
  #wireCombobox(
    dropdown: HTMLElement | undefined,
    listbox: HTMLElement | undefined,
    input: HTMLInputElement | undefined,
    id: string,
    anchor: string,
  ): void {
    if (listbox) {
      listbox.id = id;
      listbox.style.positionAnchor = anchor;
    }
    if (input) input.setAttribute("aria-controls", id);
    if (dropdown) dropdown.style.anchorName = anchor;
  }

  /** Select the option matching `value` once the combobox's control + listbox
   *  are both ready. fluent-dropdown's connectedCallback enqueues insertControl()
   *  (drops the seeded <input>, renders its own bound to an internal observable)
   *  and the listbox slots async, so selectOption() — which both marks the
   *  option selected and writes its text to the control — must wait. Mirrors
   *  <docen-ribbon-combobox>'s syncValue. */
  #syncCombobox(
    dropdown: ComboboxLike | undefined,
    listbox: HTMLElement | undefined,
    value: string,
  ): void {
    if (!dropdown || !listbox) return;
    let tries = 0;
    const apply = (): void => {
      if (!this.isConnected || tries++ > 10) return;
      if (!dropdown.listbox || !dropdown.control) {
        requestAnimationFrame(apply);
        return;
      }
      let idx = -1;
      listbox.querySelectorAll("fluent-option").forEach((opt, i) => {
        if (opt.getAttribute("value") === value) idx = i;
      });
      if (idx >= 0) dropdown.selectOption(idx);
    };
    requestAnimationFrame(apply);
  }
}

export default DocenOptionsDialog;
