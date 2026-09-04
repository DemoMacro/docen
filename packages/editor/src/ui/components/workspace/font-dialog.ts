import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import {
  FONT_NAMES,
  FONT_SIZES_CN,
  FONT_SIZES_PT,
  UNDERLINE_STYLES,
} from "../../../document/font-lists";
import { observeLang, resolveLang, t } from "../../i18n/localize";

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

/** Word's underline color dropdown (Automatic + the standard color row). */
const UNDERLINE_COLORS: ReadonlyArray<readonly [string, string]> = [
  ["000000", "color-black"],
  ["800000", "color-darkRed"],
  ["008000", "color-green"],
  ["000080", "color-darkBlue"],
  ["FF0000", "color-red"],
  ["FF00FF", "color-magenta"],
  ["FFFF00", "color-yellow"],
  ["00FFFF", "color-cyan"],
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
  select {
    min-width: 0;
    flex: 1 1 auto;
    box-sizing: border-box;
    font: inherit;
    padding: 3px 4px;
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
          <select ${ref("fontSel")}></select>
        </div>
        <div class="field">
          <label ${ref("styleLabel")}></label>
          <select ${ref("styleSel")}>
            <option value="regular"></option>
            <option value="italic"></option>
            <option value="bold"></option>
            <option value="boldItalic"></option>
          </select>
        </div>
        <div class="field">
          <label ${ref("sizeLabel")}></label>
          <select ${ref("sizeSel")}></select>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("underlineLabel")}></label>
          <select ${ref("underlineSel")}>
            <option value="">${""}</option>
          </select>
        </div>
        <div class="field">
          <label ${ref("underlineColorLabel")}></label>
          <select ${ref("underlineColorSel")}></select>
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

/** A checkbox widget plus its checked state accessor (fluent-checkbox exposes
 *  `currentChecked`, not the native `checked`). */
type FluentCheckbox = HTMLElement & { currentChecked?: boolean };

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
  @observable fontSel?: HTMLSelectElement;
  @observable styleLabel?: HTMLElement;
  @observable styleSel?: HTMLSelectElement;
  @observable sizeLabel?: HTMLElement;
  @observable sizeSel?: HTMLSelectElement;
  @observable underlineLabel?: HTMLElement;
  @observable underlineSel?: HTMLSelectElement;
  @observable underlineColorLabel?: HTMLElement;
  @observable underlineColorSel?: HTMLSelectElement;
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
    if (this.styleSel) this.styleSel.value = style;
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
    const style = this.styleSel?.value ?? "regular";
    const patch: FontDialogPatch = {
      font: this.fontSel?.value || null,
      size: this.sizeSel?.value || null,
      bold: style === "bold" || style === "boldItalic",
      italic: style === "italic" || style === "boldItalic",
      underlineStyle: this.underlineSel?.value || null,
      underlineColor: this.underlineColorSel?.value || null,
      strike: this.strike?.currentChecked ?? false,
      doubleStrike: this.doubleStrike?.currentChecked ?? false,
      superscript: this.superscript?.currentChecked ?? false,
      subscript: this.subscript?.currentChecked ?? false,
      smallCaps: this.smallCaps?.currentChecked ?? false,
      allCaps: this.allCaps?.currentChecked ?? false,
      hidden: this.hiddenCheck?.currentChecked ?? false,
    };
    this.$emit("font:ok", patch);
    this.hide();
  }

  /** Populate the font/size/underline/underline-color combos. The font and
   *  size lists are re-filled on every show() so a value outside the preset
   *  ladder (e.g. 13pt) still appears as the picked option. */
  #fillCombos(): void {
    if (!this.fontSel || !this.sizeSel || !this.underlineSel || !this.underlineColorSel) return;
    const fontValue = this.fontSel.value;
    this.fontSel.replaceChildren(
      ...FONT_NAMES.map((name) => new Option(name, name)),
      ...(fontValue && !FONT_NAMES.includes(fontValue) ? [new Option(fontValue, fontValue)] : []),
    );
    const zh = resolveLang(this).toLowerCase().startsWith("zh");
    const sizeValue = this.sizeSel.value;
    // A zh locale lists the Chinese names above the numeric points (the same
    // ladder the ribbon's size combobox shows); the emitted value is always
    // the pt string.
    const ladder = [
      ...(zh ? FONT_SIZES_CN.map(([name, pt]) => new Option(`${name} (${pt})`, String(pt))) : []),
      ...FONT_SIZES_PT.map((pt) => new Option(String(pt), String(pt))),
    ];
    this.sizeSel.replaceChildren(
      ...ladder,
      ...(sizeValue && !ladder.some((o) => o.value === sizeValue)
        ? [new Option(sizeValue, sizeValue)]
        : []),
    );
    this.underlineSel.replaceChildren(
      new Option(t("ribbon.opt.none", this), ""),
      ...UNDERLINE_STYLES.map(([value, key]) => new Option(t(`ribbon.opt.${key}`, this), value)),
    );
    this.underlineColorSel.replaceChildren(
      new Option(t("fontDialog.colorAuto", this), ""),
      ...UNDERLINE_COLORS.map(([hex, key]) => new Option(t(`fontDialog.${key}`, this), hex)),
    );
    // Re-selecting after a refill keeps the visible value stable across
    // language switches.
    if (this.fontSel.dataset.picked) this.fontSel.value = this.fontSel.dataset.picked;
  }

  /** Pick a value, adding a temporary option when it's off the preset ladder. */
  #pick(sel: HTMLSelectElement | undefined, value: string): void {
    if (!sel) return;
    if (value && !Array.from(sel.options).some((o) => o.value === value))
      sel.add(new Option(value, value), 0);
    sel.value = value;
    sel.dataset.picked = value;
  }

  #check(box: FluentCheckbox | undefined, value: boolean): void {
    if (box) box.currentChecked = value;
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
      const [regular, italic, bold, boldItalic] = this.styleSel.options;
      regular.textContent = t("fontDialog.fsRegular", this);
      italic.textContent = t("fontDialog.fsItalic", this);
      bold.textContent = t("fontDialog.fsBold", this);
      boldItalic.textContent = t("fontDialog.fsBoldItalic", this);
    }
    if (this.underlineSel) this.underlineSel.options[0].textContent = t("ribbon.opt.none", this);
    if (this.underlineColorSel)
      this.underlineColorSel.options[0].textContent = t("fontDialog.colorAuto", this);
  }
}

export default DocenFontDialog;
