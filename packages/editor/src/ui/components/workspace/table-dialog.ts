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

/** The hover grid size (Word's Insert Table grid is 10 columns × 8 rows). */
const GRID_COLS = 10;
const GRID_ROWS = 8;
const CELLS = GRID_COLS * GRID_ROWS;

/** A `fluent-text-input` widget plus its string value accessor (the value
 *  lives on the `value` property, like a native input). */
type FluentTextInput = HTMLElement & { value: string };

const cell = html`<div class="cell"></div>`;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(340px, 92vw);
  }
  .table-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .hint {
    text-align: center;
    color: #605e5c;
    min-height: 1em;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(10, 20px);
    gap: 2px;
    margin-inline: auto;
  }
  .cell {
    width: 20px;
    height: 20px;
    box-sizing: border-box;
    border: 1px solid #d1d1d1;
    background: transparent;
  }
  .cell.on {
    border-color: #6264a7;
    background: #d2d4f8;
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
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
`;

const template = html<DocenTableDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="table-body">
      <div ?hidden="${(x) => x.mode !== "grid"}">
        <div class="hint" ${ref("hint")}></div>
        <div
          class="grid"
          ${ref("gridEl")}
          @mouseover="${(x, c) => x.highlight(c.event as MouseEvent)}"
          @click="${(x, c) => x.pick(c.event as MouseEvent)}"
        >
          ${repeat(Array(CELLS).fill(0), cell)}
        </div>
        <fluent-button
          appearance="subtle"
          style="margin-top:8px;width:100%"
          @click="${(x) => x.showForm()}"
        >
          <span ${ref("formLink")}></span>
        </fluent-button>
      </div>
      <div ?hidden="${(x) => x.mode !== "form"}">
        <div class="row">
          <div class="field">
            <label ${ref("colsLabel")}></label>
            <fluent-text-input
              ${ref("colsInput")}
              type="number"
              min="1"
              max="10"
            ></fluent-text-input>
          </div>
          <div class="field">
            <label ${ref("rowsLabel")}></label>
            <fluent-text-input
              ${ref("rowsInput")}
              type="number"
              min="1"
              max="50"
            ></fluent-text-input>
          </div>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ?hidden="${(x) => x.mode !== "form"}"
        ${ref("okBtn")}
        @click="${(x) => x.applyForm()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-table-dialog>` — Word's "Insert Table" entry in both shapes: the
 * hover grid (the Table button's dropdown, where a pick inserts immediately)
 * and the classic dialog (column/row counts). `show(mode)` opens either; the
 * grid emits `table-grid:insert` with `{rows, cols}` for the host to insert
 * via the insert-table command.
 */
@customElement({ name: "docen-table-dialog", template, styles })
class DocenTableDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  /** "grid" — the hover grid; "form" — the column/row dialog. */
  @observable mode: "grid" | "form" = "grid";
  @observable hint?: HTMLElement;
  @observable gridEl?: HTMLElement;
  @observable formLink?: HTMLElement;
  @observable colsLabel?: HTMLElement;
  @observable colsInput?: FluentTextInput;
  @observable rowsLabel?: HTMLElement;
  @observable rowsInput?: FluentTextInput;
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

  /** Open the entry in one of its two shapes (Word opens the grid from the
   *  Table button and the dialog from the dropdown's "Insert Table…" item). */
  show(mode: "grid" | "form" = "grid"): void {
    this.mode = mode;
    if (mode === "form") {
      if (this.colsInput) this.colsInput.value = "3";
      if (this.rowsInput) this.rowsInput.value = "3";
    }
    this.#applyLabels();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-visible link into the classic dialog shape. */
  showForm(): void {
    this.show("form");
  }

  /** Hover the grid: light the N×M prefix (row-major from the top-left). */
  highlight(event: MouseEvent): void {
    const idx = this.#cellIndex(event);
    if (!this.gridEl) return;
    const cols = idx < 0 ? 0 : (idx % GRID_COLS) + 1;
    const rows = idx < 0 ? 0 : Math.floor(idx / GRID_COLS) + 1;
    [...this.gridEl.children].forEach((el, i) => {
      const r = Math.floor(i / GRID_COLS);
      const c = i % GRID_COLS;
      el.classList.toggle("on", r < rows && c < cols);
    });
    if (this.hint) {
      this.hint.textContent =
        idx < 0
          ? ""
          : t("tableGrid.gridHint", this).replace("{0}", String(cols)).replace("{1}", String(rows));
    }
  }

  /** Template-visible click on the grid — insert the lit shape. */
  pick(event: MouseEvent): void {
    const idx = this.#cellIndex(event);
    if (idx < 0) return;
    const cols = (idx % GRID_COLS) + 1;
    const rows = Math.floor(idx / GRID_COLS) + 1;
    this.$emit("table-grid:insert", { rows, cols });
    this.hide();
  }

  /** Template-visible OK on the classic dialog shape. */
  applyForm(): void {
    const cols = Math.max(1, Math.min(10, Math.trunc(Number(this.colsInput?.value)) || 3));
    const rows = Math.max(1, Math.min(50, Math.trunc(Number(this.rowsInput?.value)) || 3));
    this.$emit("table-grid:insert", { rows, cols });
    this.hide();
  }

  #cellIndex(event: MouseEvent): number {
    const target = (event.target as HTMLElement | null)?.closest?.(".cell");
    if (!(target instanceof HTMLElement) || !this.gridEl) return -1;
    return [...this.gridEl.children].indexOf(target);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("tableGrid.title", this);
    if (this.formLink) this.formLink.textContent = t("ribbon.opt.insert-table", this);
    if (this.colsLabel) this.colsLabel.textContent = t("tableGrid.columns", this);
    if (this.rowsLabel) this.rowsLabel.textContent = t("tableGrid.rows", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenTableDialog;
