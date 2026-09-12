import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { ModifyStylePatch } from "../../../document/extensions/commands";
import {
  CJK_FONT_NAMES,
  FONT_NAMES,
  FONT_SIZES_CN,
  FONT_SIZES_PT,
} from "../../../document/font-lists";
import { observeLang, resolveLang, t } from "../../i18n/localize";
import { appendMenuItems, renderIcon } from "../ribbon/command-helpers";
import { opt, pick, pickLadder, pickedValue, type FluentDropdown } from "./fluent-combo";

/** A paragraph style as the basedOn/next drop-downs list it. */
export interface StyleChoice {
  id: string;
  name: string;
}

/** The prefill the host passes to `show()`: the patch fields read from the
 *  style's current definition, plus the display data the dialog can't reach
 *  (the style's own name, the full paragraph-style list), the merged
 *  formatting preview CSS and the localized format-summary line. */
export interface ModifyStyleState extends ModifyStylePatch {
  name: string;
  choices: StyleChoice[];
  /** Inline CSS for the preview's sample text (the style's effective run). */
  previewCss?: string;
  /** The effective-format summary shown under the preview (host-composed). */
  description?: string;
}

const styles = css`
  :host {
    display: contents;
  }
  /* The shell reads the width via the docen-dialog channel (::part width
     rules never reach FAST's fixed-positioned native surface). 552px wraps
     the format row at its ribbon-width controls plus the dialog chrome. */
  docen-dialog {
    --dialog-width: min(552px, 94vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .section-label {
    font-weight: 600;
    opacity: 0.8;
  }
  /* The properties block: one grid, so the labels share a column and every
     control starts at the same x with the same width. The label text alone
     spreads across its run — the trailing colon rides right after it, out
     of the justification. */
  .props {
    display: grid;
    grid-template-columns: max-content 1fr;
    gap: 8px 12px;
    align-items: center;
  }
  .props > label {
    display: flex;
    white-space: nowrap;
  }
  .props > label .fill {
    flex: 1;
    text-align: justify;
    text-align-last: justify;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  /* The properties controls fill their grid column. The format row's
     comboboxes take the ribbon's fixed widths — the font box 120px and the
     size box the short 112px, same as the ribbon's font/size pair — so the
     dialog hugs its content instead of stretching one control. */
  .props > fluent-dropdown {
    width: 100%;
    min-width: 0;
  }
  .row > fluent-dropdown {
    min-width: 0;
    flex: none;
  }
  fluent-dropdown.font-sel {
    width: 120px;
  }
  fluent-dropdown.size-sel {
    width: 112px;
  }
  fluent-dropdown.scope-sel {
    flex: none;
    width: 88px;
    margin-inline-start: auto;
  }
  fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  fluent-text-input {
    flex: 1 1 auto;
    min-width: 0;
    /* FAST caps the input at 400px — the name row must fill the dialog. */
    max-width: none;
  }
  /* The comboboxes open upward — over the properties block, not over the
     format rows below (the listbox popover is anchored by FAST; the outer
     tree's rule beats the component's :host placement). */
  fluent-listbox[popover] {
    inset-block-start: auto;
    inset-block-end: anchor(top);
  }
  /* The compact format controls — subtle (borderless, the ribbon's look);
     light-DOM glyphs, so the size is owned here. */
  fluent-toggle-button.mini {
    min-width: 30px;
    min-height: 30px;
    padding: 0 4px;
    flex: none;
  }
  fluent-toggle-button.mini svg {
    display: block;
    width: 16px;
    height: 16px;
  }
  .para-row {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-wrap: wrap;
  }
  .sep {
    width: 1px;
    height: 18px;
    background: var(--neutral-stroke-rest, #d1d1d1);
    flex: none;
  }
  /* The Format trigger's chevron — slotted into the button's end slot, so
     the host's flex centering (and its column-gap) aligns it with the label
     instead of riding the text baseline one line up. */
  .fmt-caret {
    display: inline-flex;
  }
  .fmt-caret svg {
    display: block;
    width: 12px;
    height: 12px;
    fill: currentColor;
  }
  .preview {
    border: 1px solid var(--neutral-stroke-rest, #d1d1d1);
    border-radius: 4px;
    padding: 6px 10px 8px;
  }
  .preview > label {
    font-size: 12px;
    opacity: 0.7;
  }
  .preview-text {
    margin-top: 2px;
    white-space: nowrap;
    overflow: hidden;
  }
  .desc {
    font-size: 11px;
    opacity: 0.65;
    line-height: 1.4;
  }
  .checks {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
  .scope-row {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
  }
  /* The action footer: Format on the far left, Cancel/OK on the right. */
  .actions {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
  }
  .actions > .spacer {
    flex: 1;
  }
`;

const template = html<DocenModifyStyleDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="section-label" ${ref("propsLabel")}></div>
      <div class="props">
        <label ${ref("nameLabel")}></label>
        <fluent-text-input appearance="outline" ${ref("nameInput")}></fluent-text-input>
        <label ${ref("typeLabel")}></label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("typeSel")}>
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
      <div class="section-label" ${ref("formatLabel")}></div>
      <div class="row">
        <fluent-dropdown type="combobox" appearance="outline" class="font-sel" ${ref("fontSel")}>
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
        <fluent-dropdown type="combobox" appearance="outline" class="size-sel" ${ref("sizeSel")}>
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
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("boldBtn")}
          @click="${(x) => x.pressMini("bold")}"
        ></fluent-toggle-button>
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("italicBtn")}
          @click="${(x) => x.pressMini("italic")}"
        ></fluent-toggle-button>
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("underlineBtn")}
          @click="${(x) => x.pressMini("underline")}"
        ></fluent-toggle-button>
        <docen-color-picker
          icon="font-color"
          event="modify-style-color"
          no-split
          ${ref("colorPick")}
        ></docen-color-picker>
        <fluent-dropdown type="combobox" appearance="outline" class="scope-sel" ${ref("scopeSel")}>
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
      <div class="para-row" ${ref("paraRow")}></div>
      <div class="preview">
        <label ${ref("previewLabel")}></label>
        <div class="preview-text" ${ref("previewText")}></div>
      </div>
      <div class="desc" ${ref("descLabel")}></div>
      <div class="checks">
        <label class="check-field">
          <fluent-checkbox part="quick" ${ref("quickCheck")}></fluent-checkbox>
          <span ${ref("quickCheckLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox part="auto" ${ref("autoCheck")}></fluent-checkbox>
          <span ${ref("autoCheckLabel")}></span>
        </label>
      </div>
      <div class="scope-row">
        <label class="check-field">
          <fluent-radio name="modify-style-scope" ${ref("scopeDoc")}></fluent-radio>
          <span ${ref("scopeDocLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-radio name="modify-style-scope" disabled ${ref("scopeTemplate")}></fluent-radio>
          <span ${ref("scopeTemplateLabel")}></span>
        </label>
      </div>
    </div>
    <div slot="action" class="actions">
      <fluent-menu ${ref("formatMenu")}>
        <fluent-button slot="trigger" appearance="outline" ${ref("formatBtn")}>
          <span ${ref("formatBtnLabel")}></span>
        </fluent-button>
        <fluent-menu-list focusgroup="menu" ${ref("formatList")}></fluent-menu-list>
      </fluent-menu>
      <span class="spacer"></span>
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyPatch()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A checkbox/radio widget. The rendered state follows `checked`; writing
 *  `currentChecked` alone doesn't render (a user click syncs both). */
type FluentCheckbox = HTMLElement & { checked?: boolean };

/** Word's alignment single-pick group: patch value, icon, tooltip key. */
const ALIGN_GROUP: ReadonlyArray<readonly [string, string, string]> = [
  ["left", "align-left", "ribbon.cmd.align-left"],
  ["center", "align-center", "ribbon.cmd.align-center"],
  ["right", "align-right", "ribbon.cmd.align-right"],
  ["both", "justify", "ribbon.cmd.justify"],
];

/** The line-spacing single-pick group: twips-of-a-line, the three-bar glyph
 *  (wider bar gaps = the larger multiple), tooltip key. */
const LINE_GROUP: ReadonlyArray<readonly [number, string, string]> = [
  [240, "M2 4.5h12M2 8h12M2 11.5h12", "modifyStyleDialog.lineSingle"],
  [360, "M2 3.5h12M2 8h12M2 12.5h12", "modifyStyleDialog.line15"],
  [480, "M2 2.5h12M2 8h12M2 13.5h12", "modifyStyleDialog.lineDouble"],
];

/**
 * `<docen-modify-style-dialog>` — the Word "Modify Style" dialog in its
 * two-block layout: the style's properties (name / type / based on /
 * next-paragraph style) and its formatting (a font + size pair, B/I/U, a
 * color picker and a font-scope dropdown; a paragraph row of alignment and
 * line-spacing toggles), then the preview, the effective-format summary, the
 * gallery/auto-update flags and the scope radios. All controls are Fluent
 * widgets in the ribbon's subtle (borderless) style. OK emits
 * `modify-style:ok` with a full {@link ModifyStylePatch} for the host to
 * stamp through the `modify-style` command; Cancel / Esc just close. The
 * Format button opens the Font/Paragraph dialogs against the same style
 * (Tabs/Border/Numbering are honest gaps today).
 */
@customElement({ name: "docen-modify-style-dialog", template, styles })
class DocenModifyStyleDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable propsLabel?: HTMLElement;
  @observable nameLabel?: HTMLElement;
  @observable nameInput?: HTMLElement & { value?: string };
  @observable typeLabel?: HTMLElement;
  @observable typeSel?: FluentDropdown;
  @observable basedOnLabel?: HTMLElement;
  @observable basedOnSel?: FluentDropdown;
  @observable nextLabel?: HTMLElement;
  @observable nextSel?: FluentDropdown;
  @observable formatLabel?: HTMLElement;
  @observable fontSel?: FluentDropdown;
  @observable sizeSel?: FluentDropdown;
  @observable boldBtn?: HTMLElement;
  @observable italicBtn?: HTMLElement;
  @observable underlineBtn?: HTMLElement;
  @observable colorPick?: HTMLElement;
  @observable scopeSel?: FluentDropdown;
  @observable paraRow?: HTMLElement;
  @observable previewLabel?: HTMLElement;
  @observable previewText?: HTMLElement;
  @observable descLabel?: HTMLElement;
  @observable quickCheck?: FluentCheckbox;
  @observable quickCheckLabel?: HTMLElement;
  @observable autoCheck?: FluentCheckbox;
  @observable autoCheckLabel?: HTMLElement;
  @observable scopeDoc?: HTMLElement;
  @observable scopeDocLabel?: HTMLElement;
  @observable scopeTemplate?: HTMLElement;
  @observable scopeTemplateLabel?: HTMLElement;
  @observable formatMenu?: HTMLElement;
  @observable formatBtn?: HTMLElement;
  @observable formatBtnLabel?: HTMLElement;
  @observable formatList?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  /** The style being modified — re-emitted with the patch on OK. */
  #id = "";
  /** The show-time display name — a name edit only rides the patch when it
   *  actually changed (an untouched name never shadows the built-in). */
  #originalName = "";
  /** The state `show()` received — the CJK/Latin toggle re-fills from it. */
  #state?: ModifyStyleState;
  // The pending format block, edited in place and committed on OK.
  #bold = false;
  #italic = false;
  #underline = false;
  #color: string | null = null;
  #alignment: string | null = null;
  #lineSpacing: number | null = null;
  #indentLeft: number | null = null;
  #indentRight: number | null = null;
  #spacingBefore: number | null = null;
  #spacingAfter: number | null = null;
  /** The font-list scope the middle dropdown filters by (All covers both). */
  #fontScope = "all";
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    // The B/I/U glyphs (the template slots stay empty); the host owns the
    // pressed state on them, same as the paragraph-row toggles.
    for (const btn of [this.boldBtn, this.italicBtn, this.underlineBtn]) {
      if (btn) (btn as unknown as { press?: () => void }).press = () => {};
    }
    if (this.boldBtn) renderIcon(this.boldBtn, "bold");
    if (this.italicBtn) renderIcon(this.italicBtn, "italic");
    if (this.underlineBtn) renderIcon(this.underlineBtn, "underline");
    this.#bindColorPicker();
    this.#applyLabels();
    this.#fillFormatMenu();
    this.#buildParaRow();
    this.#unobserveLang = observeLang(() => {
      this.#applyLabels();
      this.#fillFormatMenu();
      this.#buildParaRow();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(state: ModifyStyleState): void {
    this.#id = state.id;
    this.#state = state;
    this.#originalName = state.name;
    this.#fontScope = "all";
    this.#bold = state.bold === true;
    this.#italic = state.italic === true;
    this.#underline = state.underline === true;
    this.#color = state.color ?? null;
    if (this.nameInput) this.nameInput.value = state.name;
    if (this.previewText) {
      this.previewText.textContent = state.name;
      this.previewText.style.cssText = state.previewCss ?? "";
    }
    if (this.descLabel) this.descLabel.textContent = state.description ?? "";
    this.#fillChoices(this.basedOnSel, state.choices, state.basedOn ?? "", true);
    this.#fillChoices(this.nextSel, state.choices, state.next ?? "", true);
    this.#fillType();
    this.#fillScopeCombo();
    this.#fillCombos();
    this.#setPressed(this.boldBtn, this.#bold);
    this.#setPressed(this.italicBtn, this.#italic);
    this.#setPressed(this.underlineBtn, this.#underline);
    if (this.colorPick) this.colorPick.setAttribute("default-color", state.color || "000000");
    this.#alignment = state.alignment ?? null;
    this.#lineSpacing = state.lineSpacing ?? null;
    this.#indentLeft = state.indentLeft ?? null;
    this.#indentRight = state.indentRight ?? null;
    this.#spacingBefore = state.spacingBefore ?? null;
    this.#spacingAfter = state.spacingAfter ?? null;
    this.#buildParaRow();
    if (this.quickCheck) this.quickCheck.checked = state.quickFormat ?? true;
    if (this.autoCheck) this.autoCheck.checked = state.autoRedefine ?? false;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible Format menu handler — opens the Font/Paragraph dialogs
   *  against this style (the host stamps their patch onto the definition). */
  applyFormat(target: "font" | "paragraph"): void {
    (this.formatMenu as unknown as { closeMenu?: () => void } | undefined)?.closeMenu?.();
    this.$emit("modify-style:format", { id: this.#id, target });
  }

  /** Template-visible scope pick — refills the font list, keeping the
   *  current pick (pickLadder re-adds off-list values). */
  onScopeChange(): void {
    const value = pickedValue(this.scopeSel);
    if (value === "cjk" || value === "latin" || value === "all") this.#fontScope = value;
    this.#fillFontCombo();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyPatch(): void {
    const name = this.nameInput?.value?.trim() ?? "";
    const patch: ModifyStylePatch = {
      id: this.#id,
      // null (the blank "(inherit)" option) commits null — inherit again.
      basedOn: pickedValue(this.basedOnSel),
      next: pickedValue(this.nextSel),
      font: pickedValue(this.fontSel),
      size: pickedValue(this.sizeSel) ? Number(pickedValue(this.sizeSel)) : null,
      bold: this.#bold,
      italic: this.#italic,
      underline: this.#underline,
      color: this.#color,
      alignment: this.#alignment,
      lineSpacing: this.#lineSpacing,
      indentLeft: this.#indentLeft,
      indentRight: this.#indentRight,
      spacingBefore: this.#spacingBefore,
      spacingAfter: this.#spacingAfter,
      quickFormat: this.quickCheck?.checked ?? undefined,
      autoRedefine: this.autoCheck?.checked ?? undefined,
    };
    if (name && name !== this.#originalName) patch.name = name;
    this.$emit("modify-style:ok", patch);
    this.hide();
  }

  /** Template-visible B/I/U press — the host flips the state and the visual
   *  itself (FAST's own flip is neutralized below; two writers would race). */
  pressMini(which: "bold" | "italic" | "underline"): void {
    const btn =
      which === "bold" ? this.boldBtn : which === "italic" ? this.italicBtn : this.underlineBtn;
    const on = !((btn as unknown as { pressed?: boolean } | undefined)?.pressed === true);
    this.#setPressed(btn, on);
    if (which === "bold") this.#bold = on;
    else if (which === "italic") this.#italic = on;
    else this.#underline = on;
  }

  /** The color picker's command — "none" clears to the automatic color; a
   *  theme pick carries a val fallback (the resolved hex), which is what the
   *  patch stores. */
  readonly #onColor = (event: Event): void => {
    const v = (event as CustomEvent).detail?.value as string | { val?: string } | undefined;
    this.#color = v == null || v === "none" ? null : typeof v === "string" ? v : (v.val ?? null);
  };

  #bindColorPicker(): void {
    if (!this.colorPick || this.colorPick.dataset.bound) return;
    this.colorPick.dataset.bound = "1";
    this.colorPick.addEventListener("command", this.#onColor);
  }

  /** The Format button's two entries (Word's list). Re-filled per language. */
  #fillFormatMenu(): void {
    if (!this.formatList) return;
    appendMenuItems(
      this.formatList,
      [
        { text: t("modifyStyleDialog.formatFont", this), value: "font" },
        { text: t("modifyStyleDialog.formatParagraph", this), value: "paragraph" },
      ],
      (item) => this.applyFormat(item.value as "font" | "paragraph"),
    );
  }

  /** The type dropdown is display-only today: every style the dialog edits is
   *  a paragraph style; a one-option list reads as fixed without needing a
   *  disabled mode on the combobox. */
  #fillType(): void {
    const listbox = this.typeSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    listbox.replaceChildren(opt(t("modifyStyleDialog.typeParagraph", this), "paragraph"));
    pick(listbox.parentElement as FluentDropdown, "paragraph");
  }

  /** The font list — the faces under the current scope (All merges both
   *  sets, CJK first). */
  #fillFontCombo(): void {
    const listbox = this.fontSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    const names =
      this.#fontScope === "cjk"
        ? CJK_FONT_NAMES
        : this.#fontScope === "latin"
          ? FONT_NAMES
          : [...new Set([...CJK_FONT_NAMES, ...FONT_NAMES])];
    listbox.replaceChildren(...names.map((name) => opt(name, name)));
    pickLadder(this.fontSel, this.#state?.font ?? "");
  }

  /** The scope dropdown's three entries; re-filled per language. */
  #fillScopeCombo(): void {
    const listbox = this.scopeSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    listbox.replaceChildren(
      opt(t("modifyStyleDialog.allFonts", this), "all"),
      opt(t("modifyStyleDialog.cjkFonts", this), "cjk"),
      opt(t("modifyStyleDialog.latinFonts", this), "latin"),
    );
    pick(this.scopeSel, this.#fontScope);
  }

  #fillCombos(): void {
    const state = this.#state;
    const listBoxes = [this.fontSel, this.sizeSel].map((s) => s?.querySelector("fluent-listbox"));
    if (listBoxes.some((b) => !b)) return;
    const [, sizeBox] = listBoxes;
    this.#fillFontCombo();
    const zh = resolveLang(this).toLowerCase().startsWith("zh");
    const size = state?.size != null ? String(state.size) : "";
    const ladder = [
      ...(zh ? FONT_SIZES_CN.map(([name, pt]) => opt(`${name} (${pt})`, String(pt))) : []),
      ...FONT_SIZES_PT.map((pt) => opt(String(pt), String(pt))),
    ];
    sizeBox!.replaceChildren(...ladder);
    pickLadder(this.sizeSel, size);
  }

  /** The basedOn/next lists: every paragraph style, headed by a blank
   *  "(inherit)" option so either pointer can be cleared. */
  #fillChoices(
    sel: FluentDropdown | undefined,
    choices: StyleChoice[],
    picked: string,
    blank: boolean,
  ): void {
    const listbox = sel?.querySelector("fluent-listbox");
    if (!listbox) return;
    const options: HTMLElement[] = [];
    if (blank) options.push(opt(t("modifyStyleDialog.inherit", this), ""));
    for (const c of choices) options.push(opt(c.name, c.id));
    if (picked && !options.some((o) => o.getAttribute("value") === picked))
      options.splice(blank ? 1 : 0, 0, opt(picked, picked));
    listbox.replaceChildren(...options);
    pick(sel, picked);
  }

  /** Write a toggle button's Fluent-owned pressed state (the show-time seed;
   *  a user click flips it directly and the click handlers read it back). */
  #setPressed(el: HTMLElement | null | undefined, on: boolean): void {
    if (el) (el as unknown as { pressed?: boolean }).pressed = on;
  }

  /** One single-pick toggle button (an icon or an inline SVG glyph). The
   *  host owns `pressed`: FAST binds its own flip on connect, i.e. *after*
   *  these listeners (bound pre-append), so its late flip would overwrite
   *  the state this dialog just recorded. */
  #iconToggle(icon: string, title: string, bars?: HTMLElement): HTMLElement {
    const b = document.createElement("fluent-toggle-button");
    b.className = "mini";
    b.setAttribute("appearance", "subtle");
    b.title = title;
    (b as unknown as { press?: () => void }).press = () => {};
    if (bars) b.append(bars);
    else renderIcon(b, icon);
    return b;
  }

  /** The three-bar line-spacing glyph — wider bar gaps read as the larger
   *  multiple (Word's ladder icons). */
  #barsSvg(path: string): HTMLElement {
    const span = document.createElement("span");
    span.innerHTML =
      `<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true">` +
      `<path d="${path}" fill="none" stroke="currentColor" stroke-width="1.3" ` +
      `stroke-linecap="round"/></svg>`;
    return span;
  }

  /** The paragraph row: alignment (single-pick icons) and line spacing
   *  (single-pick three-bar glyphs). Indents and spacing have no inline
   *  entries — they stay with the Format > Paragraph dialog. Rebuilt per
   *  show() and per language. */
  #buildParaRow(): void {
    const row = this.paraRow;
    if (!row) return;
    row.replaceChildren();
    // Alignment — one live single-pick group; the host flips `pressed` itself
    // (press() is neutralized in #iconToggle, so this is the only writer).
    for (const [value, icon, titleKey] of ALIGN_GROUP) {
      const b = this.#iconToggle(icon, t(titleKey, this));
      b.dataset.value = value;
      this.#setPressed(b, this.#alignment === value);
      b.addEventListener("click", () => {
        const on = !((b as unknown as { pressed?: boolean }).pressed === true);
        this.#setPressed(b, on);
        for (const other of ALIGN_GROUP)
          if (other[0] !== value)
            this.#setPressed(row.querySelector(`[data-value="${other[0]}"]`), false);
        this.#alignment = on ? value : null;
      });
      row.append(b);
    }
    row.append(this.#separator());
    // Line spacing — the same single-pick semantics in line multiples.
    for (const [value, path, titleKey] of LINE_GROUP) {
      const b = this.#iconToggle("", t(titleKey, this), this.#barsSvg(path));
      b.dataset.line = String(value);
      this.#setPressed(b, this.#lineSpacing === value);
      b.addEventListener("click", () => {
        const on = !((b as unknown as { pressed?: boolean }).pressed === true);
        this.#setPressed(b, on);
        for (const other of LINE_GROUP)
          if (other[0] !== value)
            this.#setPressed(row.querySelector(`[data-line="${other[0]}"]`), false);
        this.#lineSpacing = on ? value : null;
      });
      row.append(b);
    }
  }

  #separator(): HTMLElement {
    const sep = document.createElement("div");
    sep.className = "sep";
    return sep;
  }

  /** A properties-block label: the text spreads justified, the trailing
   *  colon (any language) follows it unstretched. */
  #propsLabel(el: HTMLElement | undefined, key: string): void {
    if (!el) return;
    const raw = t(key, this);
    const match = raw.match(/^(.*?)([：:])\s*$/);
    const fill = document.createElement("span");
    fill.className = "fill";
    fill.textContent = match ? match[1] : raw;
    el.replaceChildren(fill, match ? match[2] : "");
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("modifyStyleDialog.title", this);
    if (this.propsLabel) this.propsLabel.textContent = t("modifyStyleDialog.props", this);
    this.#propsLabel(this.nameLabel, "modifyStyleDialog.name");
    this.#propsLabel(this.typeLabel, "modifyStyleDialog.type");
    this.#propsLabel(this.basedOnLabel, "modifyStyleDialog.basedOn");
    this.#propsLabel(this.nextLabel, "modifyStyleDialog.next");
    if (this.formatLabel) this.formatLabel.textContent = t("modifyStyleDialog.formatSection", this);
    this.#fillScopeCombo();
    if (this.previewLabel) this.previewLabel.textContent = t("modifyStyleDialog.preview", this);
    if (this.quickCheckLabel)
      this.quickCheckLabel.textContent = t("modifyStyleDialog.quickGallery", this);
    if (this.autoCheckLabel)
      this.autoCheckLabel.textContent = t("modifyStyleDialog.autoUpdate", this);
    if (this.scopeDocLabel) this.scopeDocLabel.textContent = t("modifyStyleDialog.scopeDoc", this);
    if (this.scopeTemplateLabel)
      this.scopeTemplateLabel.textContent = t("modifyStyleDialog.scopeTemplate", this);
    if (this.formatBtnLabel) this.formatBtnLabel.textContent = t("modifyStyleDialog.format", this);
    if (this.formatBtn && !this.formatBtn.querySelector(".fmt-caret")) {
      const caret = document.createElement("span");
      caret.className = "fmt-caret";
      // The end slot makes the caret a host-level flex item — centered like
      // the label (the default slot rides .content's inline baseline).
      caret.slot = "end";
      renderIcon(caret, "caret");
      this.formatBtn.append(caret);
    }
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    for (const sel of [this.basedOnSel, this.nextSel]) {
      const blank = sel?.querySelector("fluent-option");
      if (blank) blank.textContent = t("modifyStyleDialog.inherit", this);
    }
  }
}

export default DocenModifyStyleDialog;
