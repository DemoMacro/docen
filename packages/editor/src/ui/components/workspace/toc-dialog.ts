import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** The tab leader choices (OOXML ST_TabTlc), keyed by their token. */
export const TOC_LEADERS: readonly { token: string; key: string }[] = [
  { token: "dot", key: "tocDialog.leader.dot" },
  { token: "hyphen", key: "tocDialog.leader.hyphen" },
  { token: "underscore", key: "tocDialog.leader.underscore" },
  { token: "middleDot", key: "tocDialog.leader.middleDot" },
  { token: "none", key: "tocDialog.leader.none" },
];

const leaderOptionTemplate = html<(typeof TOC_LEADERS)[number], DocenTocDialog>`
  <fluent-option value="${(l) => l.token}">${(l) => t(l.key)}</fluent-option>
`;

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
    gap: 14px;
    font-size: 13px;
  }
  fluent-field {
    align-self: flex-start;
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
`;

const template = html<DocenTocDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-field label-position="after">
        <fluent-checkbox
          slot="input"
          ${ref("pageNumbersBox")}
          @change=${(x) => x.syncGates()}
        ></fluent-checkbox>
        <label slot="label">${(x) => t("tocDialog.showPageNumbers", x)}</label>
      </fluent-field>
      <fluent-field label-position="after">
        <fluent-checkbox
          slot="input"
          ${ref("alignBox")}
          @change=${(x) => x.syncGates()}
        ></fluent-checkbox>
        <label slot="label">${(x) => t("tocDialog.alignPageNumbers", x)}</label>
      </fluent-field>
      <div class="field">
        <label>${(x) => t("tocDialog.tabLeader", x)}</label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("leaderDropdown")}>
          <fluent-listbox popover="manual" tabindex="-1">
            ${repeat(() => TOC_LEADERS, leaderOptionTemplate)}
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
        <label>${(x) => t("tocDialog.levels", x)}</label>
        <fluent-text-input ${ref("levelsInput")}></fluent-text-input>
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

/**
 * `<docen-toc-dialog>` — Word's Table of Contents dialog (References → Table
 * of Contents → Custom Table of Contents): page numbers on/off, their right
 * alignment (driving the leader tab), the tab leader glyph, and the heading
 * level window. Commits via `toc:ok` `{ headingRange, leader, showPageNumbers,
 * alignPageNumbers }`.
 */
@customElement({ name: "docen-toc-dialog", template, styles })
class DocenTocDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable pageNumbersBox?: HTMLElement & { checked: boolean };
  @observable alignBox?: HTMLElement & { checked: boolean };
  @observable leaderDropdown?: HTMLElement & { value: string | null };
  @observable levelsInput?: HTMLElement & { value: string };
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

  show(): void {
    if (this.pageNumbersBox) this.pageNumbersBox.checked = true;
    if (this.alignBox) this.alignBox.checked = true;
    if (this.leaderDropdown) this.leaderDropdown.value = "dot";
    if (this.levelsInput) this.levelsInput.value = "3";
    this.syncGates();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible OK handler (FAST templates live outside the class, so a
   *  `#`-private method can't be referenced from the binding). */
  commit(): void {
    const levels = Math.min(9, Math.max(1, Number(this.levelsInput?.value ?? 3) || 3));
    this.$emit("toc:ok", {
      headingRange: `1-${levels}`,
      leader: this.leaderDropdown?.value ?? "dot",
      showPageNumbers: this.pageNumbersBox?.checked ?? true,
      alignPageNumbers: this.alignBox?.checked ?? true,
    });
    this.hide();
  }

  /** Unaligned numbers have no tab, so the leader and the alignment itself
   *  gray out with "Right align page numbers" off (Word's linked states). */
  syncGates(): void {
    const numbers = this.pageNumbersBox?.checked ?? true;
    const align = numbers && (this.alignBox?.checked ?? true);
    if (this.alignBox) (this.alignBox as unknown as { disabled: boolean }).disabled = !numbers;
    if (this.leaderDropdown)
      (this.leaderDropdown as unknown as { disabled: boolean }).disabled = !align;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("tocDialog.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenTocDialog;
