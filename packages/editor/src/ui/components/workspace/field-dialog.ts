import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { FIELD_CATEGORIES, instructionName } from "../../../document/fields";
import { observeLang, t } from "../../i18n/localize";
import { opt, pick, pickedValue, type FluentDropdown, type FluentListbox } from "./fluent-combo";

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
    gap: 8px;
  }
  .row > label {
    min-width: 92px;
  }
  .row fluent-dropdown {
    flex: 1;
    min-width: 0;
  }
  .row fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  fluent-listbox.list {
    width: 100%;
    min-height: 168px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
    box-shadow: none;
  }
  fluent-text-input {
    flex: 1;
    min-width: 0;
  }
`;

type FluentTextInput = HTMLElement & { value: string };

const template = html<DocenFieldDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <label ${ref("categoryLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("categorySel")}
          @change="${(x) => x.syncCategory()}"
        >
          <fluent-listbox popover="manual" tabindex="-1">
            <fluent-option value="all"></fluent-option>
            <fluent-option value="date"></fluent-option>
            <fluent-option value="document"></fluent-option>
            <fluent-option value="numbering"></fluent-option>
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
      <div class="row" style="align-items:flex-start">
        <label ${ref("nameLabel")} style="padding-top:4px"></label>
        <fluent-listbox
          class="list"
          ${ref("nameList")}
          @click="${(x, c) => x.syncCodeFrom(c.event.target)}"
        ></fluent-listbox>
      </div>
      <div class="row">
        <label ${ref("codeLabel")}></label>
        <fluent-text-input ${ref("codeInput")} spellcheck="false"></fluent-text-input>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyField()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-field-dialog>` — Word's Field dialog (插入 → 域): category filter,
 * field-name list, and the field code box (the instruction, always English —
 * `DATE \@ "yyyy/M/d"`). Picking a name prefills its default instruction; a
 * hand-edited code wins on commit. Opened with `show(prefill?)` (the prefill
 * is an existing field's instruction when editing); commits via
 * `field:ok` `{ instruction }` or cancels.
 */
@customElement({ name: "docen-field-dialog", template, styles })
class DocenFieldDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable categoryLabel?: HTMLElement;
  @observable categorySel?: FluentDropdown;
  @observable nameLabel?: HTMLElement;
  @observable nameList?: FluentListbox;
  @observable codeLabel?: HTMLElement;
  @observable codeInput?: FluentTextInput;
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

  show(prefill?: string): void {
    pick(this.categorySel, "all");
    this.#renderNames();
    if (prefill != null) {
      // The edit path: highlight the field's name (Word selects it in the
      // list) and put its full instruction in the code box. Assign every
      // option — the seed pick from #renderNames must clear too.
      const name = instructionName(prefill);
      for (const o of this.nameList?.querySelectorAll("fluent-option") ?? []) {
        (o as HTMLElement & { selected?: boolean }).selected =
          instructionName(o.getAttribute("value") ?? "") === name;
      }
      if (this.codeInput) this.codeInput.value = prefill;
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Category change — rebuild the name list (all = every category). */
  syncCategory(): void {
    this.#renderNames();
  }

  /** Name-list pick — mirror the clicked option's instruction into the code
   *  box (a hand edit afterwards simply overwrites it). Read off the click
   *  target: the listbox's selectedOptions can lag one click behind while
   *  FAST re-indexes freshly rendered options (and it never fires `change`,
   *  so @click on the listbox is the only signal). */
  syncCodeFrom(target: EventTarget | null): void {
    const option = target instanceof Element ? target.closest("fluent-option") : null;
    const instruction = option?.getAttribute("value");
    if (instruction && this.codeInput) this.codeInput.value = instruction;
  }

  applyField(): void {
    const instruction = this.codeInput?.value.trim();
    if (!instruction) return;
    this.$emit("field:ok", { instruction });
    this.hide();
  }

  #renderNames(): void {
    if (!this.nameList) return;
    const category = pickedValue(this.categorySel) || "all";
    const fields = FIELD_CATEGORIES.filter((c) => category === "all" || c.key === category).flatMap(
      (c) => c.fields,
    );
    this.nameList.replaceChildren(...fields.map((f) => opt(f.name, f.instruction)));
    // A native select shows its first option preselected — a standalone
    // listbox starts with nothing picked, so pick it here.
    const first = this.nameList.querySelector("fluent-option");
    if (first) (first as HTMLElement & { selected?: boolean }).selected = true;
    // FAST indexes fresh children on an idle callback — selectedOptions is
    // still empty here, so mirror the seed straight off the option element.
    if (this.codeInput) this.codeInput.value = first?.getAttribute("value") ?? "";
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("field.title", this);
    if (this.categoryLabel) this.categoryLabel.textContent = t("field.category", this);
    if (this.nameLabel) this.nameLabel.textContent = t("field.name", this);
    if (this.codeLabel) this.codeLabel.textContent = t("field.code", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.categorySel) {
      const [all, date, doc, numbering] = this.categorySel.querySelectorAll("fluent-option");
      if (all) all.textContent = t("field.category.all", this);
      if (date) date.textContent = t("field.category.date", this);
      if (doc) doc.textContent = t("field.category.document", this);
      if (numbering) numbering.textContent = t("field.category.numbering", this);
    }
  }
}

export default DocenFieldDialog;
