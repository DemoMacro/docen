import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** The document statistics the dialog shows (Word's Word Count dialog, minus
 *  the footnote/textbox toggle — those areas aren't editable yet). */
export interface WordCountStats {
  pages: number;
  words: number;
  charsWithSpaces: number;
  charsNoSpaces: number;
  paragraphs: number;
  lines: number;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(340px, 92vw);
  }
  .wc-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
  }
  /* Word's two-column readout: label left, right-aligned count, hairline
     row rules. */
  .wc-row {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    padding: 7px 2px;
    border-block-end: 1px solid var(--colorNeutralStroke2, #e0e0e0);
  }
  .wc-row:last-child {
    border-block-end: none;
  }
  .wc-row span {
    font-size: 13px;
  }
  .wc-row b {
    font-size: 14px;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
  }
`;

const template = html<DocenWordCountDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="wc-body" ${ref("bodyEl")}></div>
    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("closeBtn")}
        @click="${(x) => x.hide()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/** Row order mirrors Word's dialog (pages → words → characters → paragraphs →
 *  lines); the value column fills from the stats JSON. */
const ROW_KEYS: readonly (keyof WordCountStats)[] = [
  "pages",
  "words",
  "charsNoSpaces",
  "charsWithSpaces",
  "paragraphs",
  "lines",
];

/**
 * `<docen-word-count-dialog>` — MS Office "Word Count" dialog. The host hands
 * over the computed statistics as a JSON `stats` attribute and calls `show()`;
 * the dialog is a read-only readout with a single Close button (Word's shape).
 */
@customElement({ name: "docen-word-count-dialog", template, styles })
class DocenWordCountDialog extends FASTElement {
  @attr stats?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable bodyEl?: HTMLElement;
  @observable closeBtn?: HTMLElement;

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

  statsChanged(): void {
    this.#renderRows();
  }

  show(): void {
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("wordCount.title", this);
    if (this.closeBtn) this.closeBtn.textContent = t("wordCount.close", this);
    this.#renderRows();
  }

  #renderRows(): void {
    if (!this.bodyEl) return;
    let stats: Partial<WordCountStats> = {};
    try {
      stats = this.stats ? (JSON.parse(this.stats) as Partial<WordCountStats>) : {};
    } catch {
      stats = {};
    }
    this.bodyEl.textContent = "";
    for (const key of ROW_KEYS) {
      const row = document.createElement("div");
      row.className = "wc-row";
      const label = document.createElement("span");
      label.textContent = t(`wordCount.${key}`, this);
      const value = document.createElement("b");
      value.textContent = String(stats[key] ?? 0);
      row.append(label, value);
      this.bodyEl.append(row);
    }
  }
}

export default DocenWordCountDialog;
