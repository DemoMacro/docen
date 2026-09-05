import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import {
  FONT_NAMES,
  FONT_SIZES_CN,
  FONT_SIZES_PT,
  UNDERLINE_STYLES,
} from "../../../document/font-lists";
import { observeLang, resolveLang, t } from "../../i18n/localize";
import { listboxOf, opt, pick, pickedValue, type FluentDropdown } from "./fluent-combo";

/**
 * What the Font dialog commits on OK — every field always present (Office
 * commits the dialog atomically): the dialog reads the selection's marks via
 * `show(patch)` and emits the same shape back through `font:ok`, so the host
 * stamps the whole run state in one pass (unchecked effect = remove it).
 * `font`/`size` null = no explicit value (inherit); `underlineStyle` null =
 * no underline; `underlineColor` null = automatic (the text color).
 */
export interface FontDialogPatch {
  font: string | null;
  /** Font size in points, as the combo's display string (e.g. "12"). */
  size: string | null;
  bold: boolean;
  italic: boolean;
  underlineStyle: string | null;
  /** w:u color, hex without "#" — null means Word's automatic. */
  underlineColor: string | null;
  strike: boolean;
  doubleStrike: boolean;
  superscript: boolean;
  subscript: boolean;
  smallCaps: boolean;
  allCaps: boolean;
  /** w:vanish — Word's Hidden effect. */
  hidden: boolean;
}

/** Word's underline color dropdown (Automatic + the standard color row) — the
 *  keys are the `fontDialog.color*` i18n suffixes. */
const UNDERLINE_COLORS: ReadonlyArray<readonly [string, string]> = [
  ["000000", "colorBlack"],
  ["800000", "colorDarkRed"],
  ["008000", "colorGreen"],
  ["000080", "colorDarkBlue"],
  ["FF0000", "colorRed"],
  ["FF00FF", "colorMagenta"],
  ["FFFF00", "colorYellow"],
  ["00FFFF", "colorCyan"],
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(420px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
    flex: 1 1 0;
    min-width: 0;
  }
  .field > label {
    white-space: nowrap;
  }
  fluent-dropdown {
    min-width: 0;
    flex: 1 1 auto;
  }
  fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  .heading {
    font-weight: 600;
  }
  .checks {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px 16px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenFontDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <div class="field">
          <label ${ref("fontLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("fontSel")}>
            <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
        </div>
        <div class="field">
          <label ${ref("styleLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("styleSel")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="regular"></fluent-option>
              <fluent-option value="italic"></fluent-option>
              <fluent-option value="bold"></fluent-option>
              <fluent-option value="boldItalic"></fluent-option>
            </fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
        </div>
        <div class="field">
          <label ${ref("sizeLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("sizeSel")}>
            <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("underlineLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("underlineSel")}>
            <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
        </div>
        <div class="field">
          <label ${ref("underlineColorLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("underlineColorSel")}>
            <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
            <input
              slot="control"
              role="combobox"
              aria-haspopup="listbox"
              type="combobox"
              size="1"
              style="width:100%;box-sizing:border-box"
            />
          </fluent-dropdown>
        </div>
      </div>
      <div class="heading" ${ref("effectsHeading")}></div>
      <div class="checks">
        <label class="check-field">
          <fluent-checkbox part="strike" ${ref("strike")}></fluent-checkbox>
          <span ${ref("strikeLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="double-strike" ${ref("doubleStrike")}></fluent-checkbox>
          <span ${ref("doubleStrikeLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="superscript" ${ref("superscript")}></fluent-checkbox>
          <span ${ref("superscriptLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="subscript" ${ref("subscript")}></fluent-checkbox>
          <span ${ref("subscriptLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="small-caps" ${ref("smallCaps")}></fluent-checkbox>
          <span ${ref("smallCapsLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="all-caps" ${ref("allCaps")}></fluent-checkbox>
          <span ${ref("allCapsLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="hidden" ${ref("hiddenCheck")}></fluent-checkbox>
          <span ${ref("hiddenLabel")}></span>
        </label>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyFont()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A checkbox widget. The rendered state follows `checked`; `currentChecked`
 *  is a separate slot only user clicks keep in sync. */
type FluentCheckbox = HTMLElement & { checked?: boolean };

/**
 * `<docen-font-dialog>` — the Word "Font" dialog (Home group's dialog-box
 * launcher): font family, style, size, underline style/color and the run
 * effects in one view. The host prefills from the selection's marks via
 * `show(patch)`; OK emits `font:ok` with the same {@link FontDialogPatch}
 * shape back for the host to stamp (Cancel / Esc just close). Rides on
 * `<docen-dialog>` for the modal shell.
 */
@customElement({ name: "docen-font-dialog", template, styles })
class DocenFontDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable fontLabel?: HTMLElement;
  @observable fontSel?: FluentDropdown;
  @observable styleLabel?: HTMLElement;
  @observable styleSel?: FluentDropdown;
  @observable sizeLabel?: HTMLElement;
  @observable sizeSel?: FluentDropdown;
  @observable underlineLabel?: HTMLElement;
  @observable underlineSel?: FluentDropdown;
  @observable underlineColorLabel?: HTMLElement;
  @observable underlineColorSel?: FluentDropdown;
  @observable effectsHeading?: HTMLElement;
  @observable strike?: FluentCheckbox;
  @observable doubleStrike?: FluentCheckbox;
  @observable superscript?: FluentCheckbox;
  @observable subscript?: FluentCheckbox;
  @observable smallCaps?: FluentCheckbox;
  @observable allCaps?: FluentCheckbox;
  @observable hiddenCheck?: FluentCheckbox;
  @observable strikeLabel?: HTMLElement;
  @observable doubleStrikeLabel?: HTMLElement;
  @observable superscriptLabel?: HTMLElement;
  @observable subscriptLabel?: HTMLElement;
  @observable smallCapsLabel?: HTMLElement;
  @observable allCapsLabel?: HTMLElement;
  @observable hiddenLabel?: HTMLElement;
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

  /** Prefill every field from the selection's run state (the host reads the
   *  marks; absent values mean "inherited" and leave the combo blank). */
  show(state: FontDialogPatch): void {
    this.#fillCombos();
    this.#pick(this.fontSel, state.font ?? "");
    const style =
      state.bold && state.italic
        ? "boldItalic"
        : state.bold
          ? "bold"
          : state.italic
            ? "italic"
            : "regular";
    pick(this.styleSel, style);
    this.#pick(this.sizeSel, state.size ?? "");
    this.#pick(this.underlineSel, state.underlineStyle ?? "");
    this.#pick(this.underlineColorSel, state.underlineColor ?? "");
    this.#check(this.strike, state.strike);
    this.#check(this.doubleStrike, state.doubleStrike);
    this.#check(this.superscript, state.superscript);
    this.#check(this.subscript, state.subscript);
    this.#check(this.smallCaps, state.smallCaps);
    this.#check(this.allCaps, state.allCaps);
    this.#check(this.hiddenCheck, state.hidden);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyFont(): void {
    const style = pickedValue(this.styleSel) ?? "regular";
    const patch: FontDialogPatch = {
      font: pickedValue(this.fontSel),
      size: pickedValue(this.sizeSel),
      bold: style === "bold" || style === "boldItalic",
      italic: style === "italic" || style === "boldItalic",
      underlineStyle: pickedValue(this.underlineSel),
      underlineColor: pickedValue(this.underlineColorSel),
      strike: this.strike?.checked ?? false,
      doubleStrike: this.doubleStrike?.checked ?? false,
      superscript: this.superscript?.checked ?? false,
      subscript: this.subscript?.checked ?? false,
      smallCaps: this.smallCaps?.checked ?? false,
      allCaps: this.allCaps?.checked ?? false,
      hidden: this.hiddenCheck?.checked ?? false,
    };
    this.$emit("font:ok", patch);
    this.hide();
  }

  /** Populate the font/size/underline/underline-color combos. The font and
   *  size lists are re-filled on every show() so a value outside the preset
   *  ladder (e.g. 13pt) still appears as the picked option. */
  #fillCombos(): void {
    const boxes = [this.fontSel, this.sizeSel, this.underlineSel, this.underlineColorSel].map(
      listboxOf,
    );
    if (boxes.some((b) => !b)) return;
    const [fontBox, sizeBox, underlineBox, colorBox] = boxes;
    const fontValue = this.fontSel?.dataset.picked ?? "";
    fontBox!.replaceChildren(
      ...FONT_NAMES.map((name) => opt(name, name)),
      ...(fontValue && !FONT_NAMES.includes(fontValue) ? [opt(fontValue, fontValue)] : []),
    );
    // Re-selecting after a refill keeps the visible value stable across
    // language switches.
    pick(this.fontSel, fontValue);
    const zh = resolveLang(this).toLowerCase().startsWith("zh");
    const sizeValue = this.sizeSel?.dataset.picked ?? "";
    // A zh locale lists the Chinese names above the numeric points (the same
    // ladder the ribbon's size combobox shows); the emitted value is always
    // the pt string.
    const ladder = [
      ...(zh ? FONT_SIZES_CN.map(([name, pt]) => opt(`${name} (${pt})`, String(pt))) : []),
      ...FONT_SIZES_PT.map((pt) => opt(String(pt), String(pt))),
    ];
    sizeBox!.replaceChildren(
      ...ladder,
      ...(sizeValue && !ladder.some((o) => o.getAttribute("value") === sizeValue)
        ? [opt(sizeValue, sizeValue)]
        : []),
    );
    pick(this.sizeSel, sizeValue);
    // Single sits between None and the variants here (Word's dialog order);
    // the ribbon ladder omits it because the split's primary action IS single.
    underlineBox!.replaceChildren(
      opt(t("ribbon.opt.none", this), ""),
      opt(t("fontDialog.underlineSingle", this), "single"),
      ...UNDERLINE_STYLES.map(([value, key]) => opt(t(`ribbon.opt.${key}`, this), value)),
    );
    pick(this.underlineSel, this.underlineSel?.dataset.picked ?? "");
    colorBox!.replaceChildren(
      opt(t("fontDialog.colorAuto", this), ""),
      ...UNDERLINE_COLORS.map(([hex, key]) => opt(t(`fontDialog.${key}`, this), hex)),
    );
    pick(this.underlineColorSel, this.underlineColorSel?.dataset.picked ?? "");
  }

  /** Pick a value, adding a temporary option when it's off the preset ladder. */
  #pick(sel: FluentDropdown | undefined, value: string): void {
    const listbox = listboxOf(sel);
    if (!sel || !listbox) return;
    if (
      value &&
      ![...listbox.querySelectorAll("fluent-option")].some((o) => o.getAttribute("value") === value)
    ) {
      listbox.prepend(opt(value, value));
    }
    pick(sel, value);
    sel.dataset.picked = value;
  }

  #check(box: FluentCheckbox | undefined, value: boolean): void {
    // `checked` is what renders — writing `currentChecked` alone leaves the
    // box visually unchecked (a user click syncs both slots, a write doesn't).
    if (box) box.checked = value;
  }

  #applyLabels(): void {
    // The combos' option labels are i18n too — refilling them keeps a language
    // switch in sync (the picked values survive the rebuild).
    this.#fillCombos();
    if (this.dialogEl) this.dialogEl.heading = t("fontDialog.title", this);
    if (this.fontLabel) this.fontLabel.textContent = t("fontDialog.font", this);
    if (this.styleLabel) this.styleLabel.textContent = t("fontDialog.style", this);
    if (this.sizeLabel) this.sizeLabel.textContent = t("fontDialog.size", this);
    if (this.underlineLabel) this.underlineLabel.textContent = t("fontDialog.underline", this);
    if (this.underlineColorLabel)
      this.underlineColorLabel.textContent = t("fontDialog.underlineColor", this);
    if (this.effectsHeading) this.effectsHeading.textContent = t("fontDialog.effects", this);
    if (this.strikeLabel) this.strikeLabel.textContent = t("fontDialog.strike", this);
    if (this.doubleStrikeLabel)
      this.doubleStrikeLabel.textContent = t("fontDialog.doubleStrike", this);
    if (this.superscriptLabel)
      this.superscriptLabel.textContent = t("fontDialog.superscript", this);
    if (this.subscriptLabel) this.subscriptLabel.textContent = t("fontDialog.subscript", this);
    if (this.smallCapsLabel) this.smallCapsLabel.textContent = t("fontDialog.smallCaps", this);
    if (this.allCapsLabel) this.allCapsLabel.textContent = t("fontDialog.allCaps", this);
    if (this.hiddenLabel) this.hiddenLabel.textContent = t("fontDialog.hidden", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.styleSel) {
      const [regular, italic, bold, boldItalic] = this.styleSel.querySelectorAll("fluent-option");
      regular.textContent = t("fontDialog.fsRegular", this);
      italic.textContent = t("fontDialog.fsItalic", this);
      bold.textContent = t("fontDialog.fsBold", this);
      boldItalic.textContent = t("fontDialog.fsBoldItalic", this);
    }
    const none = listboxOf(this.underlineSel)?.querySelector("fluent-option");
    if (none) none.textContent = t("ribbon.opt.none", this);
    const auto = listboxOf(this.underlineColorSel)?.querySelector("fluent-option");
    if (auto) auto.textContent = t("fontDialog.colorAuto", this);
  }
}

export default DocenFontDialog;
