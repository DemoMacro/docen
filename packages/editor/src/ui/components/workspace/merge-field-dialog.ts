import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";
import { pick, pickedValue, type FluentDropdown } from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(340px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .field {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .field fluent-dropdown,
  .field fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .field fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
`;

const template = html<DocenMergeFieldDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="field">
        <label>${(x) => t("mergeField.column", x)}</label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("columnSel")}>
          <fluent-listbox popover="manual" tabindex="-1" ${ref("columnList")}></fluent-listbox>
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
        <label>${(x) => t("mergeField.name", x)}</label>
        <fluent-text-input ${ref("nameInput")} spellcheck="false"></fluent-text-input>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.commit()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

type FluentTextInput = HTMLElement & { value: string };
type FluentListbox = HTMLElement & { append(...children: HTMLElement[]): void };

/**
 * `<docen-merge-field-dialog>` — Insert Merge Field: pick a recipient column
 *  from the drop-down (the data source's headers) or type a field name. The
 *  name input mirrors the drop-down pick, so either path commits the same
 *  `merge-field:ok` `{ name }`.
 */
@customElement({ name: "docen-merge-field-dialog", template, styles })
class DocenMergeFieldDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable columnSel?: FluentDropdown;
  @observable columnList?: FluentListbox;
  @observable nameInput?: FluentTextInput;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
    this.columnSel?.addEventListener("change", () => {
      const picked = pickedValue(this.columnSel);
      if (picked && this.nameInput) this.nameInput.value = picked;
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(headers: string[]): void {
    if (this.columnList) {
      this.columnList.replaceChildren(
        ...headers.map((header) => {
          const option = document.createElement("fluent-option");
          option.setAttribute("value", header);
          option.textContent = header;
          return option;
        }),
      );
    }
    pick(this.columnSel, headers[0] ?? "");
    if (this.nameInput) this.nameInput.value = headers[0] ?? "";
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  commit(): void {
    const name = this.nameInput?.value.trim();
    if (!name) return;
    this.$emit("merge-field:ok", { name });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("mergeField.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenMergeFieldDialog;
