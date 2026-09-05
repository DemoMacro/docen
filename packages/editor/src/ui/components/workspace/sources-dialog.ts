import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";
import {
  listValue,
  opt,
  pick,
  pickedValue,
  type FluentDropdown,
  type FluentListbox,
} from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(480px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  fluent-listbox.list {
    width: 100%;
    min-height: 148px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
    box-shadow: none;
  }
  .toolbar {
    display: flex;
    gap: 6px;
  }
  .form {
    display: flex;
    flex-direction: column;
    gap: 8px;
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    padding: 8px;
  }
  .form[hidden] {
    display: none;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .row > label {
    min-width: 76px;
  }
  .row fluent-dropdown,
  .row fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .row fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
`;

const template = html<DocenSourcesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-listbox class="list" ${ref("listSel")}></fluent-listbox>
      <div class="toolbar">
        <fluent-button
          appearance="accent"
          ${ref("citeBtn")}
          @click="${(x) => x.insertCitation()}"
        ></fluent-button>
        <fluent-button ${ref("newBtn")} @click="${(x) => x.toggleForm()}"></fluent-button>
        <fluent-button ${ref("deleteBtn")} @click="${(x) => x.removeSource()}"></fluent-button>
      </div>
      <div class="form" ${ref("formEl")} hidden>
        <div class="row">
          <label ${ref("typeLabel")}></label>
          <fluent-dropdown type="combobox" appearance="outline" ${ref("typeSel")}>
            <fluent-listbox popover="manual" tabindex="-1">
              <fluent-option value="Book"></fluent-option>
              <fluent-option value="JournalArticle"></fluent-option>
              <fluent-option value="ArticleInAPeriodical"></fluent-option>
              <fluent-option value="DocumentFromInternetSite"></fluent-option>
              <fluent-option value="Report"></fluent-option>
              <fluent-option value="Misc"></fluent-option>
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
        <div class="row">
          <label ${ref("tagLabel")}></label>
          <fluent-text-input ${ref("tagInput")} spellcheck="false"></fluent-text-input>
        </div>
        <div class="row">
          <label ${ref("authorLabel")}></label>
          <fluent-text-input ${ref("authorInput")} spellcheck="false"></fluent-text-input>
        </div>
        <div class="row">
          <label ${ref("titleLabel")}></label>
          <fluent-text-input ${ref("titleInput")} spellcheck="false"></fluent-text-input>
        </div>
        <div class="row">
          <label ${ref("yearLabel")}></label>
          <fluent-text-input ${ref("yearInput")} spellcheck="false"></fluent-text-input>
        </div>
        <div class="row">
          <label ${ref("publisherLabel")}></label>
          <fluent-text-input ${ref("publisherInput")} spellcheck="false"></fluent-text-input>
        </div>
        <div class="toolbar">
          <fluent-button
            appearance="accent"
            ${ref("saveBtn")}
            @click="${(x) => x.saveSource()}"
          ></fluent-button>
        </div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

type FluentTextInput = HTMLElement & { value: string };

/** A source entry as the dialog edits it — office-open's SourceTypeOptions
 *  narrowed to the fields the dialog exposes (the type stays open: the
 *  document's own sources round-trip untouched). */
export interface SourceDraft {
  tag?: string;
  sourceType?: string;
  title?: string;
  year?: string;
  publisher?: string;
  author?: { authors?: { last?: string; first?: string }[] };
}

/** "Smith, John" → { last: "Smith", first: "John" }; "John Smith" → { last:
 *  "Smith", first: "John" }; a single token (张三) keeps it whole as the last
 *  name (CJK names lead with the family name). */
export const parseAuthorName = (raw: string): { last?: string; first?: string } => {
  const s = raw.trim();
  if (s.includes(",")) {
    const [last, first] = s.split(",", 2);
    return { last: last.trim(), first: first.trim() };
  }
  const parts = s.split(/\s+/).filter(Boolean);
  if (parts.length <= 1) return { last: s };
  return { last: parts[parts.length - 1], first: parts.slice(0, -1).join(" ") };
};

/** The display name list — "Last, First" pairs joined by "; ". */
export const authorNames = (source: SourceDraft): string =>
  (source.author?.authors ?? [])
    .map((person) => [person.last, person.first].filter(Boolean).join(", "))
    .join("; ");

/** The source's list line: title — authors (year). */
export const sourceLabel = (source: SourceDraft): string => {
  const authors = authorNames(source);
  const bits = [
    source.title ?? "",
    authors ? authors : "",
    source.year ? `(${source.year})` : "",
  ].filter(Boolean);
  return bits.join(" ") || source.tag || "";
};

/**
 * `<docen-sources-dialog>` — Word's Source Manager (源管理器) over the
 * document's bibliography sources (attrs.bibliography, word/bibliography.xml):
 * the list, a new-source form (type/tag/author/title/year/publisher), delete,
 * and — in cite mode — the Insert button that places a CITATION field at the
 * caret. Opened with `show(mode, sources)`; commits via `sources:ok`
 * `{ sources }` (the full replacement list) or `citation:ok` `{ tag }`.
 */
@customElement({ name: "docen-sources-dialog", template, styles })
class DocenSourcesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable listSel?: FluentListbox;
  @observable citeBtn?: HTMLElement;
  @observable newBtn?: HTMLElement;
  @observable deleteBtn?: HTMLElement;
  @observable formEl?: HTMLElement;
  @observable typeLabel?: HTMLElement;
  @observable typeSel?: FluentDropdown;
  @observable tagLabel?: HTMLElement;
  @observable tagInput?: FluentTextInput;
  @observable authorLabel?: HTMLElement;
  @observable authorInput?: FluentTextInput;
  @observable titleLabel?: HTMLElement;
  @observable titleInput?: FluentTextInput;
  @observable yearLabel?: HTMLElement;
  @observable yearInput?: FluentTextInput;
  @observable publisherLabel?: HTMLElement;
  @observable publisherInput?: FluentTextInput;
  @observable saveBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  /** "manage" shows the editing toolbar; "cite" leads with Insert Citation. */
  #mode: "manage" | "cite" = "manage";
  #sources: SourceDraft[] = [];
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

  show(mode: "manage" | "cite", sources: SourceDraft[]): void {
    this.#mode = mode;
    this.#sources = sources.map((source) => ({ ...source }));
    if (this.formEl) this.formEl.hidden = true;
    this.#applyMode();
    this.#renderList();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  #applyMode(): void {
    // In cite mode Insert is the accent action; in manage mode New is.
    if (this.citeBtn) this.citeBtn.textContent = t("sources.cite", this);
    if (this.newBtn) this.newBtn.textContent = t("sources.new", this);
    if (this.deleteBtn) this.deleteBtn.textContent = t("sources.delete", this);
    if (this.citeBtn)
      this.citeBtn.setAttribute("appearance", this.#mode === "cite" ? "accent" : "default");
    if (this.newBtn)
      this.newBtn.setAttribute("appearance", this.#mode === "manage" ? "accent" : "default");
  }

  #renderList(): void {
    if (!this.listSel) return;
    this.listSel.replaceChildren(
      ...this.#sources.map((source, index) => opt(sourceLabel(source), String(index))),
    );
    // A native select shows its first option preselected — a standalone
    // listbox starts with nothing picked, so pick it here.
    const first = this.listSel.querySelector("fluent-option");
    if (first) (first as HTMLElement & { selected?: boolean }).selected = true;
  }

  toggleForm(): void {
    if (!this.formEl) return;
    const opening = this.formEl.hidden;
    this.formEl.hidden = !opening;
    if (opening) {
      pick(this.typeSel, "Book");
      if (this.tagInput) this.tagInput.value = "";
      if (this.authorInput) this.authorInput.value = "";
      if (this.titleInput) this.titleInput.value = "";
      if (this.yearInput) this.yearInput.value = "";
      if (this.publisherInput) this.publisherInput.value = "";
    }
  }

  saveSource(): void {
    const title = this.titleInput?.value.trim() ?? "";
    const tag = this.tagInput?.value.trim() || title.replace(/\s+/g, "");
    if (!title && !tag) return;
    const authorRaw = this.authorInput?.value ?? "";
    const source: SourceDraft = {
      ...(tag ? { tag } : {}),
      sourceType: pickedValue(this.typeSel) ?? "Misc",
      ...(title ? { title } : {}),
      ...(this.yearInput?.value.trim() ? { year: this.yearInput.value.trim() } : {}),
      ...(this.publisherInput?.value.trim() ? { publisher: this.publisherInput.value.trim() } : {}),
      ...(authorRaw.trim() ? { author: { authors: [parseAuthorName(authorRaw)] } } : {}),
    };
    this.#sources.push(source);
    this.#renderList();
    // Commit the full replacement list so the host's attrs stay the truth.
    this.#commitSources();
    this.toggleForm();
  }

  removeSource(): void {
    const index = Number(listValue(this.listSel) ?? -1);
    if (!Number.isInteger(index) || index < 0 || index >= this.#sources.length) return;
    this.#sources.splice(index, 1);
    this.#renderList();
    this.#commitSources();
  }

  insertCitation(): void {
    const index = Number(listValue(this.listSel) ?? -1);
    const source = Number.isInteger(index) ? this.#sources[index] : undefined;
    if (!source?.tag) return;
    this.$emit("citation:ok", { tag: source.tag });
    this.hide();
  }

  #commitSources(): void {
    this.$emit("sources:ok", { sources: this.#sources });
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("sources.title", this);
    this.#applyMode();
    if (this.typeLabel) this.typeLabel.textContent = t("sources.type", this);
    if (this.tagLabel) this.tagLabel.textContent = t("sources.tag", this);
    if (this.authorLabel) this.authorLabel.textContent = t("sources.author", this);
    if (this.titleLabel) this.titleLabel.textContent = t("sources.titleField", this);
    if (this.yearLabel) this.yearLabel.textContent = t("sources.year", this);
    if (this.publisherLabel) this.publisherLabel.textContent = t("sources.publisher", this);
    if (this.saveBtn) this.saveBtn.textContent = t("sources.save", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.close", this);
    if (this.typeSel) {
      const [book, journal, periodical, web, report, misc] =
        this.typeSel.querySelectorAll("fluent-option");
      if (book) book.textContent = t("sources.typeBook", this);
      if (journal) journal.textContent = t("sources.typeJournal", this);
      if (periodical) periodical.textContent = t("sources.typePeriodical", this);
      if (web) web.textContent = t("sources.typeWeb", this);
      if (report) report.textContent = t("sources.typeReport", this);
      if (misc) misc.textContent = t("sources.typeMisc", this);
    }
  }
}

export default DocenSourcesDialog;
