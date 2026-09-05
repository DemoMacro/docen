import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { ModifyStylePatch } from "../../../document/extensions/commands";
import { FONT_NAMES, FONT_SIZES_CN, FONT_SIZES_PT } from "../../../document/font-lists";
import { observeLang, resolveLang, t } from "../../i18n/localize";

/** A paragraph style as the basedOn/next drop-downs list it. */
export interface StyleChoice {
  id: string;
  name: string;
}

/** The prefill the host passes to `show()`: the patch fields read from the
 *  style's current definition, plus the display data the dialog can't reach
 *  (the style's own name and the full paragraph-style list). */
export interface ModifyStyleState extends ModifyStylePatch {
  name: string;
  choices: StyleChoice[];
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 92vw);
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
  .name {
    font-weight: 600;
    flex: 1 1 auto;
  }
  .checks {
    display: flex;
    gap: 16px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
`;

const template = html<DocenModifyStyleDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <div class="field">
          <label ${ref("nameLabel")}></label>
          <span class="name" ${ref("nameValue")}></span>
        </div>
      </div>
      <div class="row">
        <div class="field">
          <label ${ref("basedOnLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("basedOnSel")}>
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
          <label ${ref("nextLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("nextSel")}>
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
        <div class="field">
          <label ${ref("colorLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("colorSel")}>
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
      <div class="checks">
        <label class="check-field">
          <fluent-checkbox part="bold" ${ref("bold")}></fluent-checkbox>
          <span ${ref("boldLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="italic" ${ref("italic")}></fluent-checkbox>
          <span ${ref("italicLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="underline" ${ref("underline")}></fluent-checkbox>
          <span ${ref("underlineLabel")}></span>
        </label>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyPatch()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A checkbox widget. The rendered state follows `checked`; `currentChecked`
 *  is a separate slot only user clicks keep in sync. */
type FluentCheckbox = HTMLElement & { checked?: boolean };

/** A `fluent-dropdown` combobox plus its picked value (null = none). */
type FluentDropdown = HTMLElement & { value: string | null };

/** One `fluent-option` — the attr form mirrors the static template options
 *  (an absent value attr would fall back to the option's text). */
function opt(text: string, value: string): HTMLElement {
  const el = document.createElement("fluent-option");
  el.textContent = text;
  el.setAttribute("value", value);
  return el;
}

/** Programmatically pick the option carrying `value`. The FAST `value`
 *  setter ignores options it hasn't indexed yet (freshly appended ones), so
 *  set `selected` on the option itself — that syncs the property and the
 *  control input immediately. */
function pick(listbox: HTMLElement | null, value: string): void {
  if (!listbox) return;
  const option = [
    ...listbox.querySelectorAll<HTMLElement & { selected?: boolean }>("fluent-option"),
  ].find((o) => o.getAttribute("value") === value);
  if (option) option.selected = true;
}

/** Word's text-color dropdown for the dialog (automatic + the standard row). */
const TEXT_COLORS: ReadonlyArray<readonly [string, string]> = [
  ["000000", "color-black"],
  ["800000", "color-darkRed"],
  ["008000", "color-green"],
  ["000080", "color-darkBlue"],
  ["FF0000", "color-red"],
  ["FF00FF", "color-magenta"],
  ["FFFF00", "color-yellow"],
  ["00FFFF", "color-cyan"],
];

/**
 * `<docen-modify-style-dialog>` — the Word "Modify Style" dialog: the style's
 * chain pointers (based on / next-paragraph style) and the run formatting the
 * styles pane edits (font, size, color, bold/italic/underline). The host
 * prefills from the style's definition via `show(state)`; OK emits
 * `modify-style:ok` with a {@link ModifyStylePatch} for the host to stamp
 * through the `modify-style` command (Cancel / Esc just close). Rides on
 * `<docen-dialog>` for the modal shell; the drop-downs are `<fluent-dropdown>`
 * comboboxes filled dynamically per style.
 */
@customElement({ name: "docen-modify-style-dialog", template, styles })
class DocenModifyStyleDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable nameLabel?: HTMLElement;
  @observable nameValue?: HTMLElement;
  @observable basedOnLabel?: HTMLElement;
  @observable basedOnSel?: FluentDropdown;
  @observable nextLabel?: HTMLElement;
  @observable nextSel?: FluentDropdown;
  @observable fontLabel?: HTMLElement;
  @observable fontSel?: FluentDropdown;
  @observable sizeLabel?: HTMLElement;
  @observable sizeSel?: FluentDropdown;
  @observable colorLabel?: HTMLElement;
  @observable colorSel?: FluentDropdown;
  @observable bold?: FluentCheckbox;
  @observable italic?: FluentCheckbox;
  @observable underline?: FluentCheckbox;
  @observable boldLabel?: HTMLElement;
  @observable italicLabel?: HTMLElement;
  @observable underlineLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  /** The style being modified — re-emitted with the patch on OK. */
  #id = "";
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

  show(state: ModifyStyleState): void {
    this.#id = state.id;
    if (this.nameValue) this.nameValue.textContent = state.name;
    this.#fillChoices(this.basedOnSel, state.choices, state.basedOn ?? "", true);
    this.#fillChoices(this.nextSel, state.choices, state.next ?? "", true);
    this.#fillCombos(state);
    if (this.bold) this.bold.checked = state.bold;
    if (this.italic) this.italic.checked = state.italic;
    if (this.underline) this.underline.checked = state.underline;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyPatch(): void {
    const patch: ModifyStylePatch = {
      id: this.#id,
      // null (the blank "(inherit)" option) commits null — inherit again.
      basedOn: this.#picked(this.basedOnSel),
      next: this.#picked(this.nextSel),
      font: this.#picked(this.fontSel),
      size: this.#picked(this.sizeSel) ? Number(this.#picked(this.sizeSel)) : null,
      bold: this.bold?.checked ?? false,
      italic: this.italic?.checked ?? false,
      underline: this.underline?.checked ?? false,
      color: this.#picked(this.colorSel),
    };
    this.$emit("modify-style:ok", patch);
    this.hide();
  }

  /** The dropdown's picked value. A user pick syncs the FAST `value`
   *  property; a programmatic prefill may leave it "" (the control input
   *  still shows the text) — fall back to the input so prefill-then-OK
   *  round-trips. */
  #picked(dd: FluentDropdown | undefined): string | null {
    if (!dd) return null;
    const input = dd.querySelector('input[slot="control"]') as HTMLInputElement | null;
    return dd.value || input?.value.trim() || null;
  }

  /** The listbox a dropdown's options live in. */
  #listbox(sel: FluentDropdown | undefined): HTMLElement | null {
    return sel?.querySelector("fluent-listbox") ?? null;
  }

  /** The basedOn/next lists: every paragraph style, headed by a blank
   *  "(inherit)" option so either pointer can be cleared. */
  #fillChoices(
    sel: FluentDropdown | undefined,
    choices: StyleChoice[],
    picked: string,
    blank: boolean,
  ): void {
    const listbox = this.#listbox(sel);
    if (!listbox) return;
    const options: HTMLElement[] = [];
    if (blank) options.push(opt(t("modifyStyleDialog.inherit", this), ""));
    for (const c of choices) options.push(opt(c.name, c.id));
    if (picked && !options.some((o) => o.getAttribute("value") === picked))
      options.splice(blank ? 1 : 0, 0, opt(picked, picked));
    listbox.replaceChildren(...options);
    pick(listbox, picked);
  }

  #fillCombos(state: ModifyStyleState): void {
    const listBoxes = [this.fontSel, this.sizeSel, this.colorSel].map((s) => this.#listbox(s));
    if (listBoxes.some((b) => !b)) return;
    const [fontBox, sizeBox, colorBox] = listBoxes;
    const fontValue = state.font ?? "";
    fontBox!.replaceChildren(
      ...FONT_NAMES.map((name) => opt(name, name)),
      ...(fontValue && !FONT_NAMES.includes(fontValue) ? [opt(fontValue, fontValue)] : []),
    );
    pick(fontBox, fontValue);
    const zh = resolveLang(this).toLowerCase().startsWith("zh");
    const size = state.size != null ? String(state.size) : "";
    const ladder = [
      ...(zh ? FONT_SIZES_CN.map(([name, pt]) => opt(`${name} (${pt})`, String(pt))) : []),
      ...FONT_SIZES_PT.map((pt) => opt(String(pt), String(pt))),
    ];
    sizeBox!.replaceChildren(
      ...ladder,
      ...(size && !ladder.some((o) => o.getAttribute("value") === size) ? [opt(size, size)] : []),
    );
    pick(sizeBox, size);
    colorBox!.replaceChildren(
      opt(t("fontDialog.colorAuto", this), ""),
      ...TEXT_COLORS.map(([hex, key]) => opt(t(`fontDialog.${key}`, this), hex)),
    );
    pick(colorBox, state.color ?? "");
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("modifyStyleDialog.title", this);
    if (this.nameLabel) this.nameLabel.textContent = t("modifyStyleDialog.name", this);
    if (this.basedOnLabel) this.basedOnLabel.textContent = t("modifyStyleDialog.basedOn", this);
    if (this.nextLabel) this.nextLabel.textContent = t("modifyStyleDialog.next", this);
    if (this.fontLabel) this.fontLabel.textContent = t("fontDialog.font", this);
    if (this.sizeLabel) this.sizeLabel.textContent = t("fontDialog.size", this);
    if (this.colorLabel) this.colorLabel.textContent = t("modifyStyleDialog.color", this);
    if (this.boldLabel) this.boldLabel.textContent = t("fontDialog.fsBold", this);
    if (this.italicLabel) this.italicLabel.textContent = t("fontDialog.fsItalic", this);
    if (this.underlineLabel) this.underlineLabel.textContent = t("fontDialog.underline", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    for (const sel of [this.basedOnSel, this.nextSel]) {
      const blank = sel?.querySelector("fluent-option");
      if (blank) blank.textContent = t("modifyStyleDialog.inherit", this);
    }
    const auto = this.colorSel?.querySelector("fluent-option");
    if (auto) auto.textContent = t("fontDialog.colorAuto", this);
  }
}

export default DocenModifyStyleDialog;
