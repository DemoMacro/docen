import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { availableLanguages, observeLang, resolveLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
    width: 100%;
  }
  .left {
    display: flex;
    gap: 14px;
  }
  /* Right cluster — Word's zoom control: a minus / plus button flanking a
     draggable slider, then the percent. The slider is a native range input
     styled to a Fluent track + accent thumb. */
  .zoom {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .step {
    width: 18px;
    height: 18px;
    padding: 0;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 3px;
    background: transparent;
    color: var(--docen-color-text-1, #242424);
    font-size: 13px;
    line-height: 1;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .step:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .slider {
    -webkit-appearance: none;
    appearance: none;
    width: 90px;
    height: 3px;
    margin: 0;
    background: var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 2px;
    cursor: pointer;
  }
  .slider::-webkit-slider-thumb {
    -webkit-appearance: none;
    appearance: none;
    width: 11px;
    height: 11px;
    border: none;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    cursor: pointer;
  }
  .slider::-moz-range-thumb {
    width: 11px;
    height: 11px;
    border: none;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    cursor: pointer;
  }
  .pct {
    min-width: 38px;
    text-align: right;
    cursor: pointer;
    border-radius: 3px;
    padding: 1px 3px;
  }
  .pct:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  /* View shortcuts left of the zoom slider — Word's status bar carries
     Reading / Print Layout / Web Layout buttons. Only Print Layout has a
     surface today; the other two stay disabled until their projections land. */
  .views {
    display: flex;
    align-items: center;
    gap: 2px;
    margin-inline-end: 6px;
  }
  .view-btn {
    width: 22px;
    height: 22px;
    padding: 0;
    border: none;
    border-radius: 3px;
    background: transparent;
    color: var(--docen-color-text-2, #424242);
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .view-btn:hover:not(:disabled) {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .view-btn:disabled {
    color: var(--docen-color-text-3, #8a8a8a);
    opacity: 0.55;
    cursor: default;
  }
  .view-btn[aria-pressed="true"] {
    background: var(--docen-color-subtle-background-selected, #e8e8e8);
    color: var(--docen-color-accent, #0f6cbd);
  }
  /* Language indicator — sat after the word count. Plain text matching the
     surrounding status copy; a click cycles through every registered locale. */
  .lang-text {
    cursor: pointer;
    padding-inline: 2px;
  }
`;

const template = html<DocenStatusBar>`
  <span class="left">
    <span class="section" ${ref("sectionEl")}></span>
    <span class="pages" ${ref("pagesEl")}></span>
    <span class="words" ${ref("wordsEl")}></span>
    <span class="lang-text" ${ref("langBtn")}></span>
  </span>
  <span class="zoom">
    <span class="views">
      <button type="button" class="view-btn" data-view="reading" disabled>
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <path
            d="M8 3.5C6.9 2.6 5.2 2 3 2v11c2.2 0 3.9.6 5 1.5 1.1-.9 2.8-1.5 5-1.5V2c-2.2 0-3.9.6-5 1.5z"
          />
          <path d="M8 3.5v11" />
        </svg>
      </button>
      <button type="button" class="view-btn" data-view="print" aria-pressed="true">
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <path d="M3.5 1.5h6l3 3v10h-9z" />
          <path d="M9.5 1.5v3h3" />
        </svg>
      </button>
      <button type="button" class="view-btn" data-view="web" disabled>
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <circle cx="8" cy="8" r="6.5" />
          <path d="M1.5 8h13M8 1.5c-4.7 4.2-4.7 8.8 0 13 4.7-4.2 4.7-8.8 0-13z" />
        </svg>
      </button>
    </span>
    <button type="button" class="step" ${ref("outBtn")} aria-label="Zoom out">−</button>
    <input
      type="range"
      class="slider"
      min="10"
      max="500"
      step="1"
      value="100"
      ${ref("slider")}
      aria-label="Zoom level"
    />
    <button type="button" class="step" ${ref("inBtn")} aria-label="Zoom in">+</button>
    <span class="pct" ${ref("pctEl")}></span>
  </span>
`;

/**
 * `<docen-status-bar>` — Word's bottom status bar: a left cluster (caret
 * section, "Page X of Y", word count) and a right zoom control (− / slider / +
 * / percent). Numeric state arrives as attributes (`section` / `page` / `total`
 * / `words` / `zoom`); the labels are localized here. Zoom interaction emits
 * `zoom:change { zoom }` (percent, 10–500) for the host to apply.
 */
@customElement({ name: "docen-status-bar", template, styles })
class DocenStatusBar extends FASTElement {
  @attr section?: string;
  @attr page?: string;
  @attr total?: string;
  @attr words?: string;
  @attr zoom?: string;

  @observable sectionEl?: HTMLElement;
  @observable pagesEl?: HTMLElement;
  @observable wordsEl?: HTMLElement;
  @observable slider?: HTMLInputElement;
  @observable pctEl?: HTMLElement;
  @observable outBtn?: HTMLButtonElement;
  @observable inBtn?: HTMLButtonElement;
  @observable langBtn?: HTMLElement;
  #unsubscribe?: () => void;

  sectionChanged(): void {
    this.#renderSection();
  }

  pageChanged(): void {
    this.#renderPages();
  }
  totalChanged(): void {
    this.#renderPages();
  }
  wordsChanged(): void {
    this.#renderWords();
  }
  zoomChanged(): void {
    this.#renderZoom();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderAll();
    // Slider drags live; the minus / plus buttons step by 10% (Word behavior).
    this.slider?.addEventListener("input", () => this.#emit(Number(this.slider?.value ?? 100)));
    this.outBtn?.addEventListener("click", () => this.#emit(Number(this.zoom ?? 100) - 10));
    this.inBtn?.addEventListener("click", () => this.#emit(Number(this.zoom ?? 100) + 10));
    this.langBtn?.addEventListener("click", () => this.#toggleLang());
    // The percent opens the Zoom dialog (Word), and the view shortcuts select
    // a document view — the host decides which views exist (only Print Layout
    // today; the other buttons ship disabled).
    this.pctEl?.addEventListener("click", () => this.#emitOpenZoom());
    for (const btn of this.shadowRoot?.querySelectorAll<HTMLButtonElement>(".view-btn") ?? []) {
      btn.addEventListener("click", () =>
        this.dispatchEvent(
          new CustomEvent("view:select", {
            bubbles: true,
            composed: true,
            detail: { view: btn.dataset.view },
          }),
        ),
      );
    }
    this.#unsubscribe = observeLang(() => {
      this.#renderAll();
      this.#renderLang();
      this.#renderViewTitles();
    });
    this.#renderViewTitles();
  }

  /** Localized tooltips for the view shortcuts (the same Word view names the
   *  View tab uses). */
  #renderViewTitles(): void {
    for (const btn of this.shadowRoot?.querySelectorAll<HTMLButtonElement>(".view-btn") ?? []) {
      const key =
        btn.dataset.view === "reading"
          ? "ribbon.cmd.read-mode"
          : btn.dataset.view === "web"
            ? "ribbon.cmd.web-layout"
            : "ribbon.cmd.print-layout";
      btn.title = t(key, this);
    }
  }

  #emitOpenZoom(): void {
    this.dispatchEvent(
      new CustomEvent("zoom:open", {
        bubbles: true,
        composed: true,
        detail: { zoom: Number(this.zoom ?? 100) },
      }),
    );
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  #emit(zoom: number): void {
    this.dispatchEvent(
      new CustomEvent("zoom:change", {
        bubbles: true,
        composed: true,
        detail: { zoom: Math.max(10, Math.min(500, Math.round(zoom))) },
      }),
    );
  }

  #renderAll(): void {
    this.#renderSection();
    this.#renderPages();
    this.#renderWords();
    this.#renderZoom();
  }

  #renderLang(): void {
    if (!this.langBtn) return;
    // Show the current locale's display name — every registered language
    // exposes `$name` via availableLanguages(), so this tracks new locales
    // automatically (no per-language branch needed).
    const current = resolveLang(this);
    const found = availableLanguages().find((l) => l.languageTag === current);
    this.langBtn.textContent = found?.$name ?? current;
  }

  #toggleLang(): void {
    // Cycle through every registered language (en → zh-CN → fr → en …),
    // not a hard-coded zh ↔ en flip. New locales register via
    // registerTranslation / addin.localizationInfo and join the rotation
    // with no change here.
    const langs = availableLanguages();
    if (langs.length < 2) return;
    const current = resolveLang(this);
    const idx = langs.findIndex((l) => l.languageTag === current);
    const next = langs[(idx + 1) % langs.length];
    this.#emitLang(next.languageTag);
  }

  #emitLang(lang: string): void {
    this.dispatchEvent(
      new CustomEvent("lang:change", { bubbles: true, composed: true, detail: { lang } }),
    );
  }

  #renderSection(): void {
    if (this.sectionEl)
      this.sectionEl.textContent = t("status.section", this).replace(
        "{n}",
        String(Number(this.section ?? 1)),
      );
  }

  #renderPages(): void {
    if (this.pagesEl)
      this.pagesEl.textContent = t("status.page-of", this)
        .replace("{page}", String(Number(this.page || 1)))
        .replace("{total}", String(Number(this.total || 1)));
  }

  #renderWords(): void {
    if (this.wordsEl)
      this.wordsEl.textContent = t("status.words", this).replace(
        "{n}",
        String(Number(this.words ?? 0)),
      );
  }

  #renderZoom(): void {
    const z = Number(this.zoom ?? 100);
    // Sync the slider without retriggering its own input handler — only write
    // when the value drifted (keyboard / ribbon zoom changed it out of band).
    if (this.slider && Number(this.slider.value) !== z) this.slider.value = String(z);
    if (this.pctEl) this.pctEl.textContent = `${z}%`;
  }
}

export default DocenStatusBar;
