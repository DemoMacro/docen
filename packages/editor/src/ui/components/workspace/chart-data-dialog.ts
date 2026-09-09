import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { ChartDataPatch } from "../../../document/extensions/commands";
import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(560px, 92vw);
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
  .row fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .grid {
    display: grid;
    gap: 4px 6px;
    align-items: center;
    overflow: auto;
    max-height: 320px;
    padding: 2px;
  }
  .grid .head {
    font-weight: 600;
  }
  .grid fluent-text-input {
    min-width: 0;
  }
  /* The row/column tool cells — small square stealth buttons. */
  .grid fluent-button.tool {
    min-width: 24px;
    padding: 0 2px;
    font-size: 12px;
    line-height: 1;
  }
  .grid .add-category {
    grid-column: 1 / -1;
    justify-self: start;
    margin-top: 2px;
  }
`;

type FluentTextInput = HTMLElement & { value: string };
type FluentButton = HTMLElement & { textContent: string };

const template = html<DocenChartDataDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="row">
        <label ${ref("titleLabel")}></label>
        <fluent-text-input ${ref("titleInput")}></fluent-text-input>
      </div>
      <div class="grid" ${ref("grid")}></div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.applyGrid()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-chart-data-dialog>` — the chart data grid (Chart Design → Edit
 * Data; Word embeds a worksheet, the flat model edits a title + a live
 * category-row × series-column grid: values and labels in text inputs,
 * with per-column ✕ / a trailing ＋ for the series and per-row ✕ / a
 * trailing "Add category" button for the categories). Opened with
 * `show(chart)` prefilled from the chart payload; commits via `chart:ok`
 * `{ patch: ChartDataPatch }` or cancels.
 */
@customElement({ name: "docen-chart-data-dialog", template, styles })
class DocenChartDataDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable titleLabel?: HTMLElement;
  @observable titleInput?: FluentTextInput;
  @observable grid?: HTMLElement;
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

  /** Prefill from the chart payload — the series carry the columns, the
   *  categories (or the first series' value count for the category-less
   *  scatter layout) the rows. */
  show(chart: Record<string, unknown> | null): void {
    const categories = (chart?.categories as string[] | undefined) ?? [];
    const series = (chart?.series as { name?: string; values?: number[] }[] | undefined) ?? [];
    if (this.titleInput) this.titleInput.value = (chart?.title as string | undefined) ?? "";
    this.#categories = [...categories];
    this.#series = series.map((s) => ({
      name: s.name ?? "",
      values: [...(s.values ?? [])],
    }));
    if (!this.#series.length) this.#series.push({ name: "", values: [] });
    this.#render();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  applyGrid(): void {
    this.#collect();
    const patch: ChartDataPatch = {
      title: this.titleInput?.value.trim() ?? "",
      categories: this.#categories.map((c) => c.trim()),
      series: this.#series.map((s, i) => ({
        name: s.name.trim() || `Series ${i + 1}`,
        values: s.values.map((v) => (Number.isFinite(v) ? v : 0)),
      })),
    };
    this.$emit("chart:ok", { patch });
    this.hide();
  }

  /** The live grid state — mutated through #collect (before a structural
   *  change) so the user's unsaved edits survive every re-render. */
  #categories: string[] = [];
  #series: { name: string; values: number[] }[] = [];

  /** Read the inputs back into the live state. DOM order: the header name
   *  inputs first, then row-major data rows (one category followed by its
   *  series values). */
  #collect(): void {
    const grid = this.grid;
    if (!grid) return;
    const inputs = [...grid.querySelectorAll("fluent-text-input")] as FluentTextInput[];
    const count = this.#series.length;
    this.#series = this.#series.map((s, i) => ({ ...s, name: inputs[i]?.value ?? s.name }));
    const perRow = count + 1;
    const rows = Math.floor((inputs.length - count) / perRow);
    const categories: string[] = [];
    for (let r = 0; r < rows; r++) {
      const at = count + r * perRow;
      categories.push(inputs[at]!.value);
      for (let s = 0; s < count; s++) {
        const raw = inputs[at + 1 + s]?.value.trim() ?? "";
        this.#series[s]!.values[r] = raw === "" ? 0 : Number.parseFloat(raw) || 0;
      }
    }
    this.#categories = categories;
  }

  #addCategory(): void {
    this.#collect();
    this.#categories.push("");
    for (const s of this.#series) s.values.push(0);
    this.#render();
  }

  #deleteCategory(row: number): void {
    if (this.#categories.length <= 1) return;
    this.#collect();
    this.#categories.splice(row, 1);
    for (const s of this.#series) s.values.splice(row, 1);
    this.#render();
  }

  #addSeries(): void {
    this.#collect();
    this.#series.push({ name: "", values: this.#categories.map(() => 0) });
    this.#render();
  }

  #deleteSeries(index: number): void {
    if (this.#series.length <= 1) return;
    this.#collect();
    this.#series.splice(index, 1);
    this.#render();
  }

  #toolButton(label: string, title: string, onClick: () => void): FluentButton {
    const btn = document.createElement("fluent-button") as FluentButton;
    btn.classList.add("tool");
    btn.textContent = label;
    btn.setAttribute("title", title);
    btn.addEventListener("click", onClick);
    return btn;
  }

  #render(): void {
    const grid = this.grid;
    if (!grid) return;
    const count = this.#series.length;
    grid.style.gridTemplateColumns = `80px repeat(${count}, minmax(64px, 1fr)) 28px`;
    const cells: HTMLElement[] = [];
    // Column tool row — one ✕ per series, then ＋ to append one.
    cells.push(document.createElement("span"));
    for (let s = 0; s < count; s++) {
      cells.push(
        this.#toolButton("✕", t("chart-data.delete-series", this), () => this.#deleteSeries(s)),
      );
    }
    cells.push(this.#toolButton("＋", t("chart-data.add-series", this), () => this.#addSeries()));
    // Header row — the empty corner, then one name cell per series.
    const corner = document.createElement("span");
    corner.className = "head";
    cells.push(corner);
    for (let s = 0; s < count; s++) {
      const input = document.createElement("fluent-text-input") as FluentTextInput;
      input.classList.add("head");
      input.value = this.#series[s]!.name;
      cells.push(input);
    }
    cells.push(document.createElement("span"));
    // One row per category (an empty label stands in for the category-less
    // scatter layout, whose rows exist only as value slots).
    const rowCount = Math.max(this.#categories.length, ...this.#series.map((s) => s.values.length));
    for (let r = 0; r < rowCount; r++) {
      const cat = document.createElement("fluent-text-input") as FluentTextInput;
      cat.value = this.#categories[r] ?? "";
      cells.push(cat);
      for (let s = 0; s < count; s++) {
        const v = this.#series[s]!.values[r];
        const input = document.createElement("fluent-text-input") as FluentTextInput;
        input.value = v == null ? "" : String(v);
        cells.push(input);
      }
      cells.push(
        this.#toolButton("✕", t("chart-data.delete-category", this), () => this.#deleteCategory(r)),
      );
    }
    const addCategory = document.createElement("fluent-button") as FluentButton;
    addCategory.classList.add("add-category");
    addCategory.textContent = t("chart-data.add-category", this);
    addCategory.addEventListener("click", () => this.#addCategory());
    cells.push(addCategory);
    grid.replaceChildren(...cells);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("chart-data.title", this);
    if (this.titleLabel) this.titleLabel.textContent = t("chart-data.chart-title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenChartDataDialog;
