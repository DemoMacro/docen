import { FASTElement, css, customElement, html, observable, repeat } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** One pane entry — the host maps its SpellingIssue list down to what the
 *  pane shows: the surface word and its replacement candidates. */
export interface SpellingPaneEntry {
  word: string;
  suggestions: string[];
}

/** One suggestion row — clicking replaces the active misspelling with it. */
const suggestionTemplate = html<string>`
  <fluent-button
    appearance="neutral"
    class="suggestion"
    @click="${(s: string, c) => c.parent.emitReplace(s, c.event as Event)}"
    >${(s: string) => s}</fluent-button
  >
`;

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
  }
  .body {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 10px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .empty {
    color: var(--docen-color-text-2, #616161);
    padding: 14px 8px;
    text-align: center;
  }
  .counter {
    color: var(--docen-color-text-2, #616161);
  }
  .word {
    font-size: 18px;
    color: #e81123;
    /* The same red wave the canvas draws under the misspelling. */
    text-decoration: underline wavy #e81123;
    text-underline-offset: 4px;
    overflow-wrap: anywhere;
  }
  .section {
    color: var(--docen-color-text-2, #616161);
  }
  .suggestions {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .suggestion {
    text-align: start;
    padding: 6px 10px;
  }
  .actions {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
  }
  .nav {
    display: flex;
    gap: 6px;
    padding: 8px 10px;
    border-top: 1px solid var(--docen-color-divider, #e2e2e2);
  }
`;

const template = html<DocenSpellingPane>`
  <div class="body">
    ${(x) =>
      x.entries.length
        ? html`
            <span class="counter">
              ${(x) => `${x.active + 1} / ${x.entries.length} · ${x.total}`}
            </span>
            <span class="word">${(x) => x.entries[x.active]?.word ?? ""}</span>
            <span class="section">${(x) => t("spelling.suggestions", x)}</span>
            <div class="suggestions">
              ${repeat((x) => x.entries[x.active]?.suggestions ?? [], suggestionTemplate)}
            </div>
            <div class="actions">
              <fluent-button appearance="neutral" @click="${(x) => x.$emit("spelling:ignore-all")}"
                >${(x) => t("spelling.ignore-all", x)}</fluent-button
              >
              <fluent-button appearance="neutral" @click="${(x) => x.$emit("spelling:add")}"
                >${(x) => t("spelling.add", x)}</fluent-button
              >
            </div>
          `
        : html`<div class="empty">${(x) => t("spelling.empty", x)}</div>`}
  </div>
  <div class="nav">
    <fluent-button appearance="neutral" @click="${(x) => x.$emit("spelling:nav", -1)}"
      >${(x) => t("spelling.previous", x)}</fluent-button
    >
    <fluent-button appearance="neutral" @click="${(x) => x.$emit("spelling:nav", 1)}"
      >${(x) => t("spelling.next", x)}</fluent-button
    >
  </div>
`;

/**
 * `<docen-spelling-pane>` — Word's Spelling pane: the current misspelling in
 * red with its suggestions, Ignore All / Add to Dictionary, and previous /
 * next stepping. The host owns the check and the navigation; the pane reports
 * intents (`spelling:replace` / `spelling:ignore-all` / `spelling:add` /
 * `spelling:nav`) back.
 */
@customElement({ name: "docen-spelling-pane", template, styles })
class DocenSpellingPane extends FASTElement {
  /** The misspellings, in document order; `active` indexes into it. */
  @observable entries: SpellingPaneEntry[] = [];
  @observable active = -1;
  /** Total occurrences across all words (an entry may occur many times). */
  @observable total = 0;

  emitReplace(suggestion: string, event: Event): void {
    event.stopPropagation();
    this.$emit("spelling:replace", suggestion);
  }

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#unobserveLang = observeLang(() => {
      this.entries = [...this.entries];
    });
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#unobserveLang?.();
  }
}

export default DocenSpellingPane;
