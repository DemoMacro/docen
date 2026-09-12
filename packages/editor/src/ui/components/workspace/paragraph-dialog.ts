import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import {
  normalizeParagraphAlignment,
  parseMeasureTwip,
  type ParagraphDialogPatch,
} from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";
// Registration happens through the workspace index's re-export (a type-only
// import here would let esbuild elide the module and drop the side effect).
import type DocenMeasureInput from "./measure-input";

/** Line-spacing select values — multiples encode as w:line 240ths under
 *  "auto"; atLeast/exactly carry points in the value input. */
type LineSpacingChoice = "single" | "lines15" | "double" | "multiple" | "atLeast" | "exactly";

const PT_TO_TWIPS = 20;
const LINE_PER_MULTIPLE = 240;

const TA_VALUES = ["auto", "top", "center", "baseline", "bottom"];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
  }
  .para-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .tabs {
    font-size: 13px;
  }
  /* Word's two-column form grid: both columns start at fixed tracks, so the
     right column's labels line up across rows (a flex row would drift with
     the left label's width). Headings span both columns. */
  .grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    column-gap: 18px;
    row-gap: 10px;
    align-items: center;
    font-size: 13px;
  }
  .grid .para-heading {
    grid-column: 1 / -1;
  }
  /* The checkbox tabs share the grid's vertical rhythm (a bare div would
     stack the headings flush against the checkbox groups). */
  .page {
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .para-heading {
    font-weight: 600;
    margin-bottom: 4px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
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
  docen-measure-input {
    min-width: 0;
    flex: 1 1 auto;
  }
  .unit {
    white-space: nowrap;
  }
  .checks {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
  /* The Indents-and-Spacing checkboxes span both grid columns (Word stacks
     them full-width under the group they belong to). */
  .check-span {
    grid-column: 1 / -1;
  }
  .hidden {
    display: none;
  }
`;

const template = html<DocenParagraphDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="para-body">
      <fluent-tablist class="tabs" ${ref("tablist")} @change="${(x) => x.onTabChange()}">
        <fluent-tab id="para-tab-indent" ${ref("indentTabBtn")}></fluent-tab>
        <fluent-tab id="para-tab-breaks" ${ref("breaksTabBtn")}></fluent-tab>
        <fluent-tab id="para-tab-asian" ${ref("asianTabBtn")}></fluent-tab>
      </fluent-tablist>

      <div class="grid" ${ref("indentPage")}>
        <div class="para-heading" ${ref("generalHeading")}></div>
        <div class="field">
          <label ${ref("alignLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("alignDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="left"></fluent-option>
              <fluent-option value="center"></fluent-option>
              <fluent-option value="right"></fluent-option>
              <fluent-option value="both"></fluent-option>
              <fluent-option value="distribute"></fluent-option>
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
          <label ${ref("outlineLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("outlineDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="-1"></fluent-option>
              <fluent-option value="0"></fluent-option>
              <fluent-option value="1"></fluent-option>
              <fluent-option value="2"></fluent-option>
              <fluent-option value="3"></fluent-option>
              <fluent-option value="4"></fluent-option>
              <fluent-option value="5"></fluent-option>
              <fluent-option value="6"></fluent-option>
              <fluent-option value="7"></fluent-option>
              <fluent-option value="8"></fluent-option>
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
        <div class="para-heading" ${ref("indentHeading")}></div>
        <div class="field">
          <label ${ref("leftLabel")}></label>
          <docen-measure-input units="char pt mm cm in" ${ref("leftInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("specialLabel")}></label>
          <fluent-dropdown
            type="combobox"
            appearance="outline"
            ${ref("specialDropdown")}
            @change="${(x) => x.onSpecialChange()}"
          >
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="none"></fluent-option>
              <fluent-option value="firstLine"></fluent-option>
              <fluent-option value="hanging"></fluent-option>
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
          <label ${ref("rightLabel")}></label>
          <docen-measure-input units="char pt mm cm in" ${ref("rightInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("specialValLabel")}></label>
          <docen-measure-input units="char pt mm cm in" ${ref("specialVal")}></docen-measure-input>
        </div>
        <label class="check-field check-span">
          <fluent-checkbox part="mirror-indents" ${ref("mirrorIndents")}></fluent-checkbox>
          <span ${ref("mirrorIndentsLabel")}></span>
        </label>
        <label class="check-field check-span">
          <fluent-checkbox part="adjust-right" ${ref("adjustRightInd")}></fluent-checkbox>
          <span ${ref("adjustRightIndLabel")}></span>
        </label>
        <div class="para-heading" ${ref("spacingHeading")}></div>
        <div class="field">
          <label ${ref("beforeLabel")}></label>
          <docen-measure-input units="line pt mm cm in" ${ref("beforeInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("lineLabel")}></label>
          <fluent-dropdown
            type="combobox"
            appearance="outline"
            ${ref("lineDropdown")}
            @change="${(x) => x.syncSpecialEnabled()}"
          >
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="single"></fluent-option>
              <fluent-option value="lines15"></fluent-option>
              <fluent-option value="double"></fluent-option>
              <fluent-option value="multiple"></fluent-option>
              <fluent-option value="atLeast"></fluent-option>
              <fluent-option value="exactly"></fluent-option>
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
          <label ${ref("afterLabel")}></label>
          <docen-measure-input units="line pt mm cm in" ${ref("afterInput")}></docen-measure-input>
        </div>
        <div class="field">
          <label ${ref("lineValLabel")}></label>
          <docen-measure-input no-unit ${ref("lineVal")}></docen-measure-input>
          <span class="unit" ${ref("ptF")}></span>
        </div>
        <label class="check-field check-span">
          <fluent-checkbox part="contextual-spacing" ${ref("contextualSpacing")}></fluent-checkbox>
          <span ${ref("contextualSpacingLabel")}></span>
        </label>
        <label class="check-field check-span">
          <fluent-checkbox part="snap-to-grid" ${ref("snapToGrid")}></fluent-checkbox>
          <span ${ref("snapToGridLabel")}></span>
        </label>
      </div>

      <div class="page hidden" ${ref("breaksPage")}>
        <div class="para-heading" ${ref("pageBreaksHeading")}></div>
        <div class="checks">
          <label class="check-field">
            <fluent-checkbox part="widow" ${ref("widow")}></fluent-checkbox>
            <span ${ref("widowLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="keep-next" ${ref("keepNext")}></fluent-checkbox>
            <span ${ref("keepNextLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="keep-lines" ${ref("keepLines")}></fluent-checkbox>
            <span ${ref("keepLinesLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="page-break" ${ref("pageBreak")}></fluent-checkbox>
            <span ${ref("pageBreakLabel")}></span>
          </label>
        </div>
        <div class="para-heading" ${ref("exceptionsHeading")}></div>
        <div class="checks">
          <label class="check-field">
            <fluent-checkbox part="suppress-ln" ${ref("suppressLn")}></fluent-checkbox>
            <span ${ref("suppressLnLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="suppress-hyphens" ${ref("suppressHyphens")}></fluent-checkbox>
            <span ${ref("suppressHyphensLabel")}></span>
          </label>
        </div>
      </div>

      <div class="page hidden" ${ref("asianPage")}>
        <div class="para-heading" ${ref("asianWrapHeading")}></div>
        <div class="checks">
          <label class="check-field">
            <fluent-checkbox part="kinsoku" ${ref("kinsoku")}></fluent-checkbox>
            <span ${ref("kinsokuLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="word-wrap" ${ref("wordWrap")}></fluent-checkbox>
            <span ${ref("wordWrapLabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="overflow-punct" ${ref("overflowPunct")}></fluent-checkbox>
            <span ${ref("overflowPunctLabel")}></span>
          </label>
        </div>
        <div class="para-heading" ${ref("charSpacingHeading")}></div>
        <div class="checks">
          <label class="check-field">
            <fluent-checkbox part="autospace-de" ${ref("autoSpaceDE")}></fluent-checkbox>
            <span ${ref("autoSpaceDELabel")}></span>
          </label>
          <label class="check-field">
            <fluent-checkbox part="autospace-dn" ${ref("autoSpaceDN")}></fluent-checkbox>
            <span ${ref("autoSpaceDNLabel")}></span>
          </label>
        </div>
        <div class="field">
          <label ${ref("textAlignLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("taDropdown")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="auto"></fluent-option>
              <fluent-option value="top"></fluent-option>
              <fluent-option value="center"></fluent-option>
              <fluent-option value="baseline"></fluent-option>
              <fluent-option value="bottom"></fluent-option>
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
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("defaultBtn")} @click="${(x) => x.applyAsDefault()}"></fluent-button>
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyParagraph()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** A checkbox widget. The rendered state follows `checked`; `currentChecked`
 *  is a separate slot only user clicks keep in sync. */
type FluentCheckbox = HTMLElement & { checked?: boolean };

/** A `fluent-dropdown` combobox plus its picked value (null = none). */
type FluentDropdown = HTMLElement & { value: string | null };

/**
 * `<docen-paragraph-dialog>` — Word's Paragraph dialog with its three tabs:
 * Indents and Spacing (alignment, outline level, indents, spacing), Line and
 * Page Breaks (pagination + formatting exceptions), and Asian Typography
 * (kinsoku/word-wrap/overflow punctuation, Asian auto-spacing, vertical text
 * alignment). The host prefills from the caret paragraph's attrs via
 * `show(attrs)`; OK emits `paragraph:ok` with a full {@link ParagraphDialogPatch}
 * for the host to stamp onto every selected paragraph (`paragraph-dialog-apply`);
 * Set As Default emits the same patch as `paragraph:default` (the host writes
 * it onto the Normal style instead). Cancel / Esc just close. Rides on
 * `<docen-dialog>` for the modal shell; all
 * drop-downs are `<fluent-dropdown>` comboboxes (the fixed option lists).
 */
@customElement({ name: "docen-paragraph-dialog", template, styles })
class DocenParagraphDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable tablist?: HTMLElement & { activeid: string };
  @observable indentTabBtn?: HTMLElement;
  @observable breaksTabBtn?: HTMLElement;
  @observable asianTabBtn?: HTMLElement;
  @observable indentPage?: HTMLElement;
  @observable breaksPage?: HTMLElement;
  @observable asianPage?: HTMLElement;
  @observable generalHeading?: HTMLElement;
  @observable alignLabel?: HTMLElement;
  @observable alignDropdown?: FluentDropdown;
  @observable outlineLabel?: HTMLElement;
  @observable outlineDropdown?: FluentDropdown;
  @observable indentHeading?: HTMLElement;
  @observable leftLabel?: HTMLElement;
  @observable leftInput?: DocenMeasureInput;
  @observable rightLabel?: HTMLElement;
  @observable rightInput?: DocenMeasureInput;
  @observable specialLabel?: HTMLElement;
  @observable specialDropdown?: FluentDropdown;
  @observable specialValLabel?: HTMLElement;
  @observable specialVal?: DocenMeasureInput;
  @observable spacingHeading?: HTMLElement;
  @observable beforeLabel?: HTMLElement;
  @observable beforeInput?: DocenMeasureInput;
  @observable afterLabel?: HTMLElement;
  @observable afterInput?: DocenMeasureInput;
  @observable lineLabel?: HTMLElement;
  @observable lineDropdown?: FluentDropdown;
  @observable lineValLabel?: HTMLElement;
  @observable lineVal?: DocenMeasureInput;
  @observable mirrorIndents?: FluentCheckbox;
  @observable mirrorIndentsLabel?: HTMLElement;
  @observable adjustRightInd?: FluentCheckbox;
  @observable adjustRightIndLabel?: HTMLElement;
  @observable snapToGrid?: FluentCheckbox;
  @observable snapToGridLabel?: HTMLElement;
  @observable contextualSpacing?: FluentCheckbox;
  @observable contextualSpacingLabel?: HTMLElement;
  @observable pageBreaksHeading?: HTMLElement;
  @observable widow?: FluentCheckbox;
  @observable keepNext?: FluentCheckbox;
  @observable keepLines?: FluentCheckbox;
  @observable pageBreak?: FluentCheckbox;
  @observable widowLabel?: HTMLElement;
  @observable keepNextLabel?: HTMLElement;
  @observable keepLinesLabel?: HTMLElement;
  @observable pageBreakLabel?: HTMLElement;
  @observable exceptionsHeading?: HTMLElement;
  @observable suppressLn?: FluentCheckbox;
  @observable suppressHyphens?: FluentCheckbox;
  @observable suppressLnLabel?: HTMLElement;
  @observable suppressHyphensLabel?: HTMLElement;
  @observable asianWrapHeading?: HTMLElement;
  @observable kinsoku?: FluentCheckbox;
  @observable wordWrap?: FluentCheckbox;
  @observable overflowPunct?: FluentCheckbox;
  @observable kinsokuLabel?: HTMLElement;
  @observable wordWrapLabel?: HTMLElement;
  @observable overflowPunctLabel?: HTMLElement;
  @observable charSpacingHeading?: HTMLElement;
  @observable autoSpaceDE?: FluentCheckbox;
  @observable autoSpaceDN?: FluentCheckbox;
  @observable autoSpaceDELabel?: HTMLElement;
  @observable autoSpaceDNLabel?: HTMLElement;
  @observable textAlignLabel?: HTMLElement;
  @observable taDropdown?: FluentDropdown;
  @observable defaultBtn?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  @observable ptF?: HTMLElement;

  #unobserveLang?: () => void;
  #tab: "indent" | "breaks" | "asian" = "indent";

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

  /** Prefill every field from the caret paragraph's attrs (verbatim PM
   *  mirror of ParagraphPropertiesOptionsBase). Absent values fall back to
   *  Word's defaults (left alignment, single spacing, widow control on,
   *  kinsoku/overflow/auto-spacing on). `styleId` opens the dialog in the
   *  Modify Style dialog's Format > Paragraph mode: the fields prefill from
   *  that style's definition and OK retargets the style (the host reads the
   *  `data-for-style` marker on `paragraph:ok`) — Word hides Set As Default
   *  there, the commit itself already is the style write. */
  show(attrs: Record<string, unknown> = {}, opts?: { styleId?: string }): void {
    if (opts?.styleId) this.dataset.forStyle = opts.styleId;
    else delete this.dataset.forStyle;
    if (this.defaultBtn) this.defaultBtn.style.display = opts?.styleId ? "none" : "";
    const indent = (attrs.indent ?? {}) as {
      left?: number;
      leftChars?: number;
      right?: number;
      rightChars?: number;
      firstLine?: number;
      firstLineChars?: number;
      hanging?: number;
      hangingChars?: number;
    };
    const spacing = (attrs.spacing ?? {}) as {
      before?: number;
      beforeLines?: number;
      after?: number;
      afterLines?: number;
      line?: number;
      lineRule?: string;
    };
    if (this.alignDropdown) this.alignDropdown.value = normalizeParagraphAlignment(attrs.alignment);
    if (this.outlineDropdown) {
      const level = typeof attrs.outlineLevel === "number" ? attrs.outlineLevel : -1;
      this.outlineDropdown.value = String(Math.max(-1, Math.min(8, level)));
    }
    this.#setMeasure(this.leftInput, indent.left, indent.leftChars, "char");
    this.#setMeasure(this.rightInput, indent.right, indent.rightChars, "char");
    let special = "none";
    if (indent.firstLine || indent.firstLineChars) special = "firstLine";
    else if (indent.hanging || indent.hangingChars) special = "hanging";
    if (this.specialDropdown) this.specialDropdown.value = special;
    this.#setMeasure(
      this.specialVal,
      special === "firstLine"
        ? indent.firstLine
        : special === "hanging"
          ? indent.hanging
          : undefined,
      special === "firstLine"
        ? indent.firstLineChars
        : special === "hanging"
          ? indent.hangingChars
          : undefined,
      "char",
    );
    this.syncSpecialEnabled();
    this.#setMeasure(this.beforeInput, spacing.before, spacing.beforeLines, "line");
    this.#setMeasure(this.afterInput, spacing.after, spacing.afterLines, "line");
    this.#prefillLine(spacing);
    this.#check(this.mirrorIndents, attrs.mirrorIndents, false);
    this.#check(this.adjustRightInd, attrs.adjustRightInd, true);
    this.#check(this.snapToGrid, attrs.snapToGrid, true);
    this.#check(this.contextualSpacing, attrs.contextualSpacing, false);
    this.#check(this.widow, attrs.widowControl, true);
    this.#check(this.keepNext, attrs.keepNext, false);
    this.#check(this.keepLines, attrs.keepLines, false);
    this.#check(this.pageBreak, attrs.pageBreakBefore, false);
    this.#check(this.suppressLn, attrs.suppressLineNumbers, false);
    this.#check(this.suppressHyphens, attrs.suppressAutoHyphens, false);
    this.#check(this.kinsoku, attrs.kinsoku, true);
    this.#check(this.wordWrap, attrs.wordWrap, false);
    this.#check(this.overflowPunct, attrs.overflowPunctuation, true);
    this.#check(this.autoSpaceDE, attrs.autoSpaceDE, true);
    this.#check(this.autoSpaceDN, attrs.autoSpaceEastAsianText, true);
    const ta = (attrs.textAlignment as string) ?? "auto";
    if (this.taDropdown) this.taDropdown.value = TA_VALUES.includes(ta) ? ta : "auto";
    this.showTab("indent");
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  onTabChange(): void {
    const id = String(this.tablist?.activeid ?? "");
    const tab = id.replace("para-tab-", "") as "indent" | "breaks" | "asian";
    // showTab also stamps activeid (programmatic opens); skip the echo.
    if (tab !== this.#tab) this.showTab(tab);
  }

  showTab(tab: "indent" | "breaks" | "asian"): void {
    this.#tab = tab;
    this.indentPage?.classList.toggle("hidden", tab !== "indent");
    this.breaksPage?.classList.toggle("hidden", tab !== "breaks");
    this.asianPage?.classList.toggle("hidden", tab !== "asian");
    if (this.tablist) this.tablist.activeid = `para-tab-${tab}`;
  }

  /** The value input only applies to multiple/atLeast/exactly spacing; the
   *  unit chip follows the rule (pt for the fixed heights, "lines" for
   *  multiples — Word swaps the unit with the rule). */
  syncSpecialEnabled(): void {
    const choice = (this.lineDropdown?.value ?? "single") as LineSpacingChoice;
    if (this.lineVal)
      this.lineVal.disabled = choice === "single" || choice === "lines15" || choice === "double";
    if (this.ptF) {
      this.ptF.textContent =
        choice === "atLeast" || choice === "exactly"
          ? t("paragraph.pt", this)
          : choice === "multiple"
            ? t("paragraph.unitMultiple", this)
            : "";
    }
    if (this.specialVal && this.specialDropdown)
      this.specialVal.disabled = this.specialDropdown.value === "none";
  }

  /** Word pre-fills the value when the special indent first leaves "none",
   *  so the box commits a real indent instead of the pre-filled zero. */
  onSpecialChange(): void {
    this.syncSpecialEnabled();
    const choice = this.specialDropdown?.value;
    const box = this.specialVal;
    if ((choice === "firstLine" || choice === "hanging") && box && !box.value) {
      box.value = 2;
      box.unit = "char";
    }
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  applyParagraph(): void {
    const patch = this.#buildPatch();
    if (!patch) return;
    this.$emit("paragraph:ok", patch);
    this.hide();
  }

  /** Set As Default — the same patch the OK button commits, but the host
   *  stamps it onto the Normal style instead of the selection. */
  applyAsDefault(): void {
    const patch = this.#buildPatch();
    if (!patch) return;
    this.$emit("paragraph:default", patch);
    this.hide();
  }

  #buildPatch(): ParagraphDialogPatch {
    // Comboboxes allow free typing, so an unmatched value falls back to the
    // Word default rather than stamping an unknown token.
    const alignment = String(this.alignDropdown?.value ?? "left");
    const ta = String(this.taDropdown?.value ?? "auto");
    // "Body Text" (-1) commits null — w:outlineLvl has no -1; absence IS
    // body text. Level 0 is a legal level and must survive.
    const outline = Number(this.outlineDropdown?.value ?? "-1");
    const leftParts = this.#commitParts(this.leftInput);
    const rightParts = this.#commitParts(this.rightInput);
    const beforeParts = this.#commitParts(this.beforeInput);
    const afterParts = this.#commitParts(this.afterInput);
    const patch: ParagraphDialogPatch = {
      alignment: normalizeParagraphAlignment(alignment),
      outlineLevel: outline >= 0 ? outline : null,
      indent: {
        left: leftParts?.twips,
        leftChars: leftParts?.hundredths,
        right: rightParts?.twips,
        rightChars: rightParts?.hundredths,
        // firstLine and hanging are mutually exclusive — the unchosen pair
        // clears (OOXML rejects both on one paragraph).
        firstLine: undefined,
        firstLineChars: undefined,
        hanging: undefined,
        hangingChars: undefined,
      },
      spacing: {
        before: beforeParts?.twips,
        beforeLines: beforeParts?.hundredths,
        after: afterParts?.twips,
        afterLines: afterParts?.hundredths,
      },
      mirrorIndents: this.mirrorIndents?.checked ?? false,
      adjustRightInd: this.adjustRightInd?.checked ?? true,
      snapToGrid: this.snapToGrid?.checked ?? true,
      contextualSpacing: this.contextualSpacing?.checked ?? false,
      widowControl: this.widow?.checked ?? true,
      keepNext: this.keepNext?.checked ?? false,
      keepLines: this.keepLines?.checked ?? false,
      pageBreakBefore: this.pageBreak?.checked ?? false,
      suppressLineNumbers: this.suppressLn?.checked ?? false,
      suppressAutoHyphens: this.suppressHyphens?.checked ?? false,
      kinsoku: this.kinsoku?.checked ?? true,
      wordWrap: this.wordWrap?.checked ?? false,
      overflowPunct: this.overflowPunct?.checked ?? true,
      autoSpaceDE: this.autoSpaceDE?.checked ?? true,
      autoSpaceDN: this.autoSpaceDN?.checked ?? true,
      textAlignment: TA_VALUES.includes(ta) ? ta : "auto",
    };
    const special = this.specialDropdown?.value;
    const specialParts = this.#commitParts(this.specialVal);
    if (special === "firstLine") {
      patch.indent.firstLine = specialParts?.twips;
      patch.indent.firstLineChars = specialParts?.hundredths;
    } else if (special === "hanging") {
      patch.indent.hanging = specialParts?.twips;
      patch.indent.hangingChars = specialParts?.hundredths;
    }
    const choice = (this.lineDropdown?.value ?? "single") as LineSpacingChoice;
    // The line rule writes its own keys; the before/after slots above stay.
    if (choice === "single") {
      patch.spacing.line = 240;
      patch.spacing.lineRule = "auto";
    } else if (choice === "lines15") {
      patch.spacing.line = 360;
      patch.spacing.lineRule = "auto";
    } else if (choice === "double") {
      patch.spacing.line = 480;
      patch.spacing.lineRule = "auto";
    } else if (choice === "multiple") {
      patch.spacing.line = Math.round((this.lineVal?.value ?? 1) * LINE_PER_MULTIPLE);
      patch.spacing.lineRule = "auto";
    } else if (choice === "atLeast") {
      patch.spacing.line = Math.round((this.lineVal?.value ?? 0) * PT_TO_TWIPS);
      patch.spacing.lineRule = "atLeast";
    } else {
      patch.spacing.line = Math.round((this.lineVal?.value ?? 0) * PT_TO_TWIPS);
      patch.spacing.lineRule = "exact";
    }
    return patch;
  }

  /** Prefills a measure box: the char/line twin wins (the count is what the
   *  user typed — Word shows the same), else the twip value in points, else
   *  the box's zero with its default unit. */
  #setMeasure(
    box: DocenMeasureInput | undefined,
    twips: number | undefined,
    hundredths: number | undefined,
    defaultUnit: "char" | "line",
  ): void {
    if (!box) return;
    if (typeof hundredths === "number") {
      box.value = hundredths / 100;
      box.unit = defaultUnit;
    } else if (typeof twips === "number") {
      box.value = twips / PT_TO_TWIPS;
      box.unit = "pt";
    } else {
      box.value = 0;
      box.unit = defaultUnit;
    }
  }

  /** A box's committed value pair: char/line units ride their own OOXML twins
   *  (hundredths — a font-relative count, not a length), every other unit
   *  resolves through the shared measure table into twips. Null = blank box. */
  #commitParts(box: DocenMeasureInput | undefined): { twips?: number; hundredths?: number } | null {
    const v = box?.value;
    if (box == null || v == null) return null;
    const unit = box.unit;
    if (unit === "char" || unit === "line") return { hundredths: Math.round(v * 100) };
    const twips = parseMeasureTwip(`${v}${unit}`);
    return { twips: Math.round(twips ?? v * PT_TO_TWIPS) };
  }

  /** Twips → points, rounded to 2 decimals for the input. */
  #pt(twips?: number): string {
    if (typeof twips !== "number") return "";
    return String(Math.round((twips / PT_TO_TWIPS) * 100) / 100);
  }

  #prefillLine(spacing: { line?: number; lineRule?: string }): void {
    if (!this.lineDropdown) return;
    const line = spacing.line ?? LINE_PER_MULTIPLE;
    const rule = spacing.lineRule ?? "auto";
    let choice: LineSpacingChoice = "single";
    let val = "";
    if (rule === "atLeast" || rule === "exact") {
      // OOXML's token is "exact"; the UI choice spells it "exactly".
      choice = rule === "atLeast" ? "atLeast" : "exactly";
      val = this.#pt(line);
    } else {
      const mult = line / LINE_PER_MULTIPLE;
      if (Math.abs(mult - 1) < 0.01) choice = "single";
      else if (Math.abs(mult - 1.5) < 0.01) choice = "lines15";
      else if (Math.abs(mult - 2) < 0.01) choice = "double";
      else {
        choice = "multiple";
        val = String(Math.round(mult * 100) / 100);
      }
    }
    this.lineDropdown.value = choice;
    if (this.lineVal) this.lineVal.value = val === "" ? null : Number(val);
    this.syncSpecialEnabled();
  }

  #check(box: FluentCheckbox | undefined, value: unknown, fallback: boolean): void {
    // `checked` is what renders — writing `currentChecked` alone leaves the
    // box visually unchecked (a user click syncs both slots, a write doesn't).
    if (box) box.checked = typeof value === "boolean" ? value : fallback;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("paragraph.title", this);
    if (this.indentTabBtn) this.indentTabBtn.textContent = t("paragraph.tabIndent", this);
    if (this.breaksTabBtn) this.breaksTabBtn.textContent = t("paragraph.tabBreaks", this);
    if (this.asianTabBtn) this.asianTabBtn.textContent = t("paragraph.tabAsian", this);
    if (this.generalHeading) this.generalHeading.textContent = t("paragraph.generalHeading", this);
    if (this.alignLabel) this.alignLabel.textContent = t("paragraph.alignment", this);
    if (this.outlineLabel) this.outlineLabel.textContent = t("paragraph.outline", this);
    if (this.indentHeading) this.indentHeading.textContent = t("paragraph.indentHeading", this);
    if (this.leftLabel) this.leftLabel.textContent = t("paragraph.left", this);
    if (this.rightLabel) this.rightLabel.textContent = t("paragraph.right", this);
    if (this.specialLabel) this.specialLabel.textContent = t("paragraph.special", this);
    if (this.specialValLabel) this.specialValLabel.textContent = t("paragraph.indentValue", this);
    if (this.spacingHeading) this.spacingHeading.textContent = t("paragraph.spacingHeading", this);
    if (this.beforeLabel) this.beforeLabel.textContent = t("paragraph.before", this);
    if (this.afterLabel) this.afterLabel.textContent = t("paragraph.after", this);
    if (this.lineLabel) this.lineLabel.textContent = t("paragraph.lineSpacing", this);
    if (this.lineValLabel) this.lineValLabel.textContent = t("paragraph.lineValue", this);
    if (this.mirrorIndentsLabel)
      this.mirrorIndentsLabel.textContent = t("paragraph.mirrorIndents", this);
    if (this.adjustRightIndLabel)
      this.adjustRightIndLabel.textContent = t("paragraph.adjustRightInd", this);
    if (this.snapToGridLabel) this.snapToGridLabel.textContent = t("paragraph.snapToGrid", this);
    if (this.contextualSpacingLabel)
      this.contextualSpacingLabel.textContent = t("paragraph.contextualSpacing", this);
    if (this.defaultBtn) this.defaultBtn.textContent = t("paragraph.setDefault", this);
    if (this.pageBreaksHeading)
      this.pageBreaksHeading.textContent = t("paragraph.pageBreaksHeading", this);
    if (this.widowLabel) this.widowLabel.textContent = t("paragraph.widowControl", this);
    if (this.keepNextLabel) this.keepNextLabel.textContent = t("paragraph.keepNext", this);
    if (this.keepLinesLabel) this.keepLinesLabel.textContent = t("paragraph.keepLines", this);
    if (this.pageBreakLabel) this.pageBreakLabel.textContent = t("paragraph.pageBreakBefore", this);
    if (this.exceptionsHeading)
      this.exceptionsHeading.textContent = t("paragraph.exceptionsHeading", this);
    if (this.suppressLnLabel)
      this.suppressLnLabel.textContent = t("paragraph.suppressLineNumbers", this);
    if (this.suppressHyphensLabel)
      this.suppressHyphensLabel.textContent = t("paragraph.suppressAutoHyphens", this);
    if (this.asianWrapHeading)
      this.asianWrapHeading.textContent = t("paragraph.asianWrapHeading", this);
    if (this.kinsokuLabel) this.kinsokuLabel.textContent = t("paragraph.kinsoku", this);
    if (this.wordWrapLabel) this.wordWrapLabel.textContent = t("paragraph.wordWrap", this);
    if (this.overflowPunctLabel)
      this.overflowPunctLabel.textContent = t("paragraph.overflowPunct", this);
    if (this.charSpacingHeading)
      this.charSpacingHeading.textContent = t("paragraph.charSpacingHeading", this);
    if (this.autoSpaceDELabel) this.autoSpaceDELabel.textContent = t("paragraph.autoSpaceDE", this);
    if (this.autoSpaceDNLabel) this.autoSpaceDNLabel.textContent = t("paragraph.autoSpaceDN", this);
    if (this.textAlignLabel) this.textAlignLabel.textContent = t("paragraph.textAlignLabel", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    // Drop-down option labels (alignment reuses the ribbon's entries).
    this.#labelOptions(this.alignDropdown, [
      t("ribbon.cmd.align-left", this),
      t("ribbon.cmd.align-center", this),
      t("ribbon.cmd.align-right", this),
      t("ribbon.cmd.justify", this),
      t("paragraph.alignDistribute", this),
    ]);
    this.#labelOptions(this.outlineDropdown, [
      t("paragraph.outlineBody", this),
      ...[1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => `${t("paragraph.level", this)} ${i}`),
    ]);
    this.#labelOptions(this.specialDropdown, [
      t("paragraph.specialNone", this),
      t("paragraph.specialFirstLine", this),
      t("paragraph.specialHanging", this),
    ]);
    this.#labelOptions(this.lineDropdown, [
      t("paragraph.lsSingle", this),
      t("paragraph.ls15", this),
      t("paragraph.lsDouble", this),
      t("paragraph.lsMultiple", this),
      t("paragraph.lsAtLeast", this),
      t("paragraph.lsExactly", this),
    ]);
    this.#labelOptions(this.taDropdown, [
      t("paragraph.taAuto", this),
      t("paragraph.taTop", this),
      t("paragraph.taCenter", this),
      t("paragraph.taBaseline", this),
      t("paragraph.taBottom", this),
    ]);
  }

  #labelOptions(dropdown: FluentDropdown | undefined, labels: string[]): void {
    if (!dropdown) return;
    const options = dropdown.querySelectorAll("fluent-option");
    options.forEach((opt, i) => {
      if (labels[i]) opt.textContent = labels[i];
    });
  }
}

export default DocenParagraphDialog;
