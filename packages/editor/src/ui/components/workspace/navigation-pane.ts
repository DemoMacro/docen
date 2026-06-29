import { observeLang, t } from "../../i18n/localize";

const template = document.createElement("template");
template.innerHTML = `
  <style>
    :host {
      display: flex;
      flex-direction: column;
      flex: 1;
      min-height: 0;
      box-sizing: border-box;
      font-size: 12px;
    }
    .search {
      padding: 8px;
      box-sizing: border-box;
      flex: 0 0 auto;
    }
    fluent-text-input {
      width: 100%;
      box-sizing: border-box;
    }
    fluent-tablist {
      flex: 0 0 auto;
      box-sizing: border-box;
      padding: 0 8px;
      border-block-end: 1px solid var(--docen-color-divider, #e2e2e2);
    }
    .content {
      flex: 1;
      min-height: 0;
      overflow: auto;
    }
    /* Show only the active tab's slot; default to Headings. */
    :host(:not([tab])) slot[name="pages"],
    :host(:not([tab])) slot[name="results"],
    :host([tab="headings"]) slot[name="pages"],
    :host([tab="headings"]) slot[name="results"],
    :host([tab="pages"]) slot[name="headings"],
    :host([tab="pages"]) slot[name="results"],
    :host([tab="results"]) slot[name="headings"],
    :host([tab="results"]) slot[name="pages"] {
      display: none;
    }
  </style>
  <div class="search" part="search">
    <fluent-text-input part="search-input" placeholder="Search">
      <svg slot="start" width="16" height="16" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">
        <path d="M6.5 2a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9ZM1 6.5a5.5 5.5 0 0 1 9.9-3.26l3.4-3.4a.75.75 0 0 1 1.06 1.06l-3.4 3.4A5.5 5.5 0 0 1 6.5 12 5.5 5.5 0 0 1 1 6.5Z" />
      </svg>
    </fluent-text-input>
  </div>
  <fluent-tablist part="tabs" activeid="nav-headings" size="small">
    <fluent-tab id="nav-headings" data-i18n="nav.headings">Headings</fluent-tab>
    <fluent-tab id="nav-pages" data-i18n="nav.pages">Pages</fluent-tab>
    <fluent-tab id="nav-results" data-i18n="nav.results">Results</fluent-tab>
  </fluent-tablist>
  <div class="content" part="content">
    <slot name="headings"></slot>
    <slot name="pages"></slot>
    <slot name="results"></slot>
  </div>`;

/**
 * `<docen-navigation-pane tab="headings">` — a Word-style Navigation Pane: a
 * search box, a Headings / Pages / Results tablist, and three named slots
 * (`headings` / `pages` / `results`) for each tab's content. Only the active
 * tab's slot is visible (CSS keyed on `host[tab]`; default Headings). Typing in
 * the search box emits `navigation:search { query }`; switching tabs emits
 * `navigation:tab { tab }`. Content-agnostic — the editor package fills the
 * slots (typically `<docen-outline slot="headings">`).
 */
class DocenNavigationPane extends HTMLElement {
  static get observedAttributes(): string[] {
    return ["tab"];
  }

  #tabs?: HTMLElement;
  #tabObserver?: MutationObserver;
  #unsubscribe?: () => void;

  attributeChangedCallback(name: string, _old: string, next: string): void {
    if (name === "tab" && next) this.#syncTabs(next);
  }

  connectedCallback(): void {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" }).append(template.content.cloneNode(true));
    }
    const root = this.shadowRoot!;
    this.#tabs = root.querySelector("[part='tabs']")!;
    this.#syncTabs(this.getAttribute("tab") ?? "headings");
    const searchInput = root.querySelector<HTMLElement>("[part='search-input']");
    searchInput?.addEventListener("input", (event: Event) => {
      const query = (event.target as HTMLInputElement).value;
      this.dispatchEvent(
        new CustomEvent("navigation:search", {
          bubbles: true,
          composed: true,
          detail: { query, source: this },
        }),
      );
    });
    // Enter = next match, Shift+Enter = previous (Word's Find behavior).
    searchInput?.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      this.dispatchEvent(
        new CustomEvent("navigation:find", {
          bubbles: true,
          composed: true,
          detail: { direction: event.shiftKey ? "prev" : "next", source: this },
        }),
      );
    });
    // fluent-tablist updates `activeid` on click; mirror it to host[tab] and
    // emit. (Its change-event detail is unreliable across versions.)
    this.#tabObserver = new MutationObserver(() => {
      const id = this.#tabs?.getAttribute("activeid") ?? "";
      const tab = id.replace(/^nav-/, "") || "headings";
      if (this.getAttribute("tab") !== tab) {
        this.setAttribute("tab", tab);
        this.dispatchEvent(
          new CustomEvent("navigation:tab", {
            bubbles: true,
            composed: true,
            detail: { tab, source: this },
          }),
        );
      }
    });
    this.#tabObserver.observe(this.#tabs, { attributes: true, attributeFilter: ["activeid"] });
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#tabObserver?.disconnect();
    this.#unsubscribe?.();
  }

  #syncTabs(tab: string): void {
    this.#tabs?.setAttribute("activeid", `nav-${tab}`);
  }

  #applyI18n(): void {
    const root = this.shadowRoot!;
    root.querySelector("[part='search-input']")?.setAttribute("placeholder", t("nav.search", this));
    root.querySelector('[data-i18n="nav.headings"]')?.replaceChildren(t("nav.headings", this));
    root.querySelector('[data-i18n="nav.pages"]')?.replaceChildren(t("nav.pages", this));
    root.querySelector('[data-i18n="nav.results"]')?.replaceChildren(t("nav.results", this));
  }
}

customElements.define("docen-navigation-pane", DocenNavigationPane);

export default DocenNavigationPane;
