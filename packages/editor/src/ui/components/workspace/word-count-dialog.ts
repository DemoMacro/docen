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
  .wc-row:last-of-type {
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
  .wc-notes {
    padding: 10px 2px 0;
  }
`;

const template = html<DocenWordCountDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="wc-body" ${ref("bodyEl")}>
      <fluent-checkbox
        class="wc-notes"
        ${ref("notesEl")}
        @change="${(x) => x.toggleNotes()}"
      ></fluent-checkbox>
    </div>
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
 * over the computed statistics as JSON — `stats` for the body and
 * `stats-extra` with textboxes/notes folded back in — and calls `show()`; the
 * checkbox (Word's, default ON) picks which readout renders.
 */
@customElement({ name: "docen-word-count-dialog", template, styles })
class DocenWordCountDialog extends FASTElement {
  @attr stats?: string;

  @attr statsExtra?: string;

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable bodyEl?: HTMLElement;
  @observable closeBtn?: HTMLElement;
  @observable notesEl?: HTMLElement & { textContent?: string; checked?: boolean };

  // Word's checkbox ships checked.
  #includeNotes = true;

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

  statsExtraChanged(): void {
    this.#renderRows();
  }

  show(): void {
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  toggleNotes(): void {
    this.#includeNotes = !this.#includeNotes;
    this.#renderRows();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("wordCount.title", this);
    if (this.closeBtn) this.closeBtn.textContent = t("wordCount.close", this);
    if (this.notesEl) {
      this.notesEl.textContent = t("wordCount.includeNotes", this);
      this.notesEl.checked = this.#includeNotes;
    }
    this.#renderRows();
  }

  #renderRows(): void {
    if (!this.bodyEl) return;
    const raw = this.#includeNotes ? (this.statsExtra ?? this.stats) : this.stats;
    let stats: Partial<WordCountStats> = {};
    try {
      stats = raw ? (JSON.parse(raw) as Partial<WordCountStats>) : {};
    } catch {
      stats = {};
    }
    for (const row of this.bodyEl.querySelectorAll(".wc-row")) row.remove();
    const notesEl = this.notesEl;
    for (const key of ROW_KEYS) {
      const row = document.createElement("div");
      row.className = "wc-row";
      const label = document.createElement("span");
      label.textContent = t(`wordCount.${key}`, this);
      const value = document.createElement("b");
      value.textContent = String(stats[key] ?? 0);
      row.append(label, value);
      // The checkbox lives inside the body container; rows insert before it
      // so the readout keeps ROW_KEYS order (pages first, lines last).
      this.bodyEl.insertBefore(row, notesEl ?? null);
    }
  }
}

export default DocenWordCountDialog;
