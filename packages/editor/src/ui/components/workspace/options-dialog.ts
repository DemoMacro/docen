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

/** The "Document" section's seed/commit shape — the host converts between
 *  twips and cm, the dialog speaks human units only. */
export interface DocumentSettings {
  /** Default tab stop in cm — undefined leaves the document's value untouched. */
  defaultTabStop?: number;
  /** Update fields when the document opens. */
  updateFields?: boolean;
  /** office-open protection token ("none" = unrestricted). */
  protection?: string;
  /** Word compatibility mode version (15 = 2013+, 14 = 2010, 12 = 2007). */
  compatVersion?: number;
}

// Per-instance CSS anchor names so each dropdown's listbox popover floats
// under its own control — without it, Fluent's default strands the popover at
// the viewport corner (same race as <docen-ribbon-combobox>).
let seq = 0;

// Protection tokens in listbox order — #applyLabels pairs them with i18n.
const protectionKeys = ["none", "readOnly", "trackedChanges", "comments", "forms"];
// OOXML version → release year, for the compat labels' i18n keys.
const compatYears: Record<string, string> = { "15": "2013", "14": "2010", "12": "2007" };

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
  .opt-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .opt-row label {
    flex: none;
    min-width: 96px;
  }
  .opt-row > *:not(label) {
    flex: 1;
    min-width: 0;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
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
      <div class="opt-field">
        <div class="opt-heading" ${ref("spellHeadingEl")}></div>
        <!-- fluent-checkbox has no default label slot (indicator slots only) —
             the label span sits outside, the wrapping <label> routes clicks. -->
        <label class="check-field">
          <fluent-checkbox ${ref("spellBox")}></fluent-checkbox>
          <span ${ref("spellLabelEl")}></span>
        </label>
      </div>
      <div class="opt-field">
        <div class="opt-heading" ${ref("markdownHeadingEl")}></div>
        <label class="check-field">
          <fluent-checkbox ${ref("markdownBox")}></fluent-checkbox>
          <span ${ref("markdownLabelEl")}></span>
        </label>
      </div>
      <div class="opt-field">
        <div class="opt-heading" ${ref("docHeadingEl")}></div>
        <div class="opt-row">
          <label ${ref("tabLabelEl")}></label>
          <fluent-text-input
            ${ref("tabInput")}
            type="number"
            step="any"
            min="0"
          ></fluent-text-input>
        </div>
        <label class="check-field">
          <fluent-checkbox ${ref("updateFieldsBox")}></fluent-checkbox>
          <span ${ref("updateFieldsLabelEl")}></span>
        </label>
        <div class="opt-row">
          <label ${ref("protectionLabelEl")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("protectionDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1" ${ref("protectionListbox")}>
              <fluent-option value="none"></fluent-option>
              <fluent-option value="readOnly"></fluent-option>
              <fluent-option value="trackedChanges"></fluent-option>
              <fluent-option value="comments"></fluent-option>
              <fluent-option value="forms"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
        <div class="opt-row">
          <label ${ref("compatLabelEl")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("compatDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1" ${ref("compatListbox")}>
              <fluent-option value="15"></fluent-option>
              <fluent-option value="14"></fluent-option>
              <fluent-option value="12"></fluent-option>
            </fluent-listbox>
            <input slot="control" role="combobox" aria-readonly="true" readonly />
          </fluent-dropdown>
        </div>
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
 * `<docen-options-dialog locale="…" theme="…" proofing="…">` — MS Office
 * "Options" dialog. v1 carries the host-level prefs: UI language (over
 * {@link availableLanguages},
 * so a locale added via `registerTranslation` or an add-in's `localizationInfo`
 * appears here with no further wiring), theme (the built-in Fluent web /
 * teams / high-contrast themes, plus any registered via registerTheme), and
 * the spell-as-you-type toggle (Proofing). The "Document" section edits the
 * open document's settings.xml-level options (tab stop / field update /
 * editing restrictions / compatibility mode) seeded via the `document`
 * property. Rides
 * on `<docen-dialog>` for
 * the modal shell (backdrop / Esc / show).
 *
 * The host seeds the current values via `locale` / `theme` / `proofing` /
 * `markdown` / `document`, calls `show()`, and listens for
 * `options:ok { lang, theme, spellcheck, markdown, document }` (确定).
 * Cancel / Esc just close.
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
  /** Whether spell-as-you-type runs ("true"/"false") — the Proofing section's
   *  checkbox pre-fills from it. Absent = on. */
  @attr proofing?: string;
  /** Whether the Markdown input mode is on ("true"/"false") — the Markdown
   *  section's checkbox pre-fills from it. Absent = on. */
  @attr markdown?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable headingEl?: HTMLElement;
  @observable dropdown?: HTMLElement;
  @observable listbox?: HTMLElement;
  @observable input?: HTMLInputElement;
  @observable themeHeadingEl?: HTMLElement;
  @observable themeDropdown?: HTMLElement;
  @observable themeListbox?: HTMLElement;
  @observable themeInput?: HTMLInputElement;
  @observable spellHeadingEl?: HTMLElement;
  @observable spellBox?: HTMLElement & { checked?: boolean };
  @observable spellLabelEl?: HTMLElement;
  @observable markdownHeadingEl?: HTMLElement;
  @observable markdownBox?: HTMLElement & { checked?: boolean };
  @observable markdownLabelEl?: HTMLElement;
  /** The document settings seed — the host assigns it before show(). */
  @observable document?: DocumentSettings;
  @observable docHeadingEl?: HTMLElement;
  @observable tabLabelEl?: HTMLElement;
  @observable tabInput?: HTMLInputElement & { value: string };
  @observable updateFieldsBox?: HTMLElement & { checked?: boolean };
  @observable updateFieldsLabelEl?: HTMLElement;
  @observable protectionLabelEl?: HTMLElement;
  @observable protectionDropdown?: HTMLElement & { value: string };
  @observable protectionListbox?: HTMLElement;
  @observable compatLabelEl?: HTMLElement;
  @observable compatDropdown?: HTMLElement & { value: string };
  @observable compatListbox?: HTMLElement;
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
    // The checkbox writes `checked` directly — programmatic currentChecked
    // never renders (the prefill/read-back rule).
    if (this.spellBox) this.spellBox.checked = this.proofing !== "false";
    if (this.markdownBox) this.markdownBox.checked = this.markdown !== "false";
    const d = this.document;
    if (this.tabInput)
      this.tabInput.value = d?.defaultTabStop != null ? String(d.defaultTabStop) : "";
    if (this.updateFieldsBox) this.updateFieldsBox.checked = d?.updateFields === true;
    if (this.protectionDropdown) this.protectionDropdown.value = d?.protection ?? "none";
    if (this.compatDropdown) this.compatDropdown.value = String(d?.compatVersion ?? 15);
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
    const tab = this.tabInput ? Number.parseFloat(this.tabInput.value) : Number.NaN;
    this.dispatchEvent(
      new CustomEvent("options:ok", {
        bubbles: true,
        composed: true,
        detail: {
          lang: this.#langLocal,
          theme: this.#themeLocal,
          spellcheck: this.spellBox?.checked !== false,
          markdown: this.markdownBox?.checked !== false,
          document: {
            defaultTabStop: Number.isFinite(tab) ? tab : undefined,
            updateFields: this.updateFieldsBox?.checked === true,
            protection: this.protectionDropdown?.value ?? "none",
            compatVersion: Number(this.compatDropdown?.value ?? 15),
          } satisfies DocumentSettings,
        },
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
    if (this.headingEl) this.headingEl.textContent = t("options.uiLanguage", this);
    if (this.themeHeadingEl) this.themeHeadingEl.textContent = t("options.theme", this);
    if (this.spellHeadingEl) this.spellHeadingEl.textContent = t("options.proofing", this);
    if (this.spellLabelEl) this.spellLabelEl.textContent = t("options.spellAsYouType", this);
    if (this.markdownHeadingEl) this.markdownHeadingEl.textContent = t("options.markdown", this);
    if (this.markdownLabelEl) this.markdownLabelEl.textContent = t("options.markdownInput", this);
    if (this.docHeadingEl) this.docHeadingEl.textContent = t("options.document", this);
    if (this.tabLabelEl) this.tabLabelEl.textContent = t("options.defaultTabStop", this);
    if (this.updateFieldsLabelEl)
      this.updateFieldsLabelEl.textContent = t("options.updateFields", this);
    if (this.protectionLabelEl) this.protectionLabelEl.textContent = t("options.protection", this);
    if (this.compatLabelEl) this.compatLabelEl.textContent = t("options.compat", this);
    this.protectionListbox?.querySelectorAll("fluent-option").forEach((opt, i) => {
      opt.textContent = t(`options.protection.${protectionKeys[i]}`, this);
    });
    // The compatibility labels key off the release year (readable i18n); the
    // option values stay the OOXML version numbers.
    this.compatListbox?.querySelectorAll("fluent-option").forEach((opt) => {
      const v = opt.getAttribute("value");
      const year = v ? compatYears[v] : undefined;
      if (v && year) opt.textContent = t(`options.compat.${year}`, this);
    });
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
