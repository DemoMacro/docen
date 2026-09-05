import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** One collected item — the text preview the pane shows plus the docen slice
 *  payload when the copy came from inside the editor (marks survive a paste
 *  from here only through it). */
export interface ClipboardEntry {
  text: string;
  payload: string | null;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
  }
  .toolbar {
    display: flex;
    gap: 6px;
    padding: 8px;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
  }
  .list {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 6px;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .empty {
    color: var(--docen-color-text-2, #616161);
    padding: 14px 8px;
    text-align: center;
  }
  .item {
    padding: 8px;
    border-radius: 4px;
    cursor: pointer;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    /* Three lines of preview — Word's pane shows a tall snippet per item. */
    display: -webkit-box;
    -webkit-line-clamp: 3;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }
  .item:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
`;

/** One collected item. Clicks are delegated from the list container — a
 *  per-item @click binding would stretch this past the formatter's line
 *  width, and the wrapped-out text node (`.item` is pre-wrap) renders the
 *  template's own indentation as blank lines above the text. A `repeat` (not
 *  a binding returning `.map()`): FAST stringifies an array a nested binding
 *  returns, rendering each template fragment as "[object Object]". */
const itemTemplate = html<ClipboardEntry>`<div class="item">${(entry) => entry.text}</div>`;

const template = html<DocenClipboardPane>`
  <div class="toolbar">
    <fluent-button appearance="neutral" @click="${(x) => x.pasteAll()}"
      >${(x) => t("clipboard.paste-all", x)}</fluent-button
    >
    <fluent-button appearance="neutral" @click="${(x) => x.clear()}"
      >${(x) => t("clipboard.clear", x)}</fluent-button
    >
  </div>
  <div class="list" ${ref("listEl")} @click="${(x, c) => x.onListClick(c.event)}">
    ${(x) =>
      x.entries.length
        ? html`${repeat((x) => x.entries, itemTemplate)}`
        : html`<div class="empty">${(x) => t("clipboard.empty", x)}</div>`}
  </div>
`;

/**
 * `<docen-clipboard-pane>` — the Office Clipboard: the session's collected
 * in-editor copies/cuts, newest first; clicking an item pastes it at the
 * caret. The host owns the collection (`clipboard:collect` pushes in) and
 * the paste actions (`clipboard:paste` / `clipboard:paste-all` come back).
 */
@customElement({ name: "docen-clipboard-pane", template, styles })
class DocenClipboardPane extends FASTElement {
  @observable listEl?: HTMLElement;
  /** Newest first; the host caps the length (Word keeps 24). */
  @observable entries: ClipboardEntry[] = [];
  /** Set once after a paste so the pane closes like Word's (options allow
   *  keeping it open — the host drives visibility). */
  @attr({ mode: "boolean" }) closeAfterPaste = false;

  emitPaste(entry: ClipboardEntry, event: Event): void {
    event.stopPropagation();
    this.$emit("clipboard:paste", entry);
  }

  /** Template-bound (public — the FAST template cannot reach a #private
   *  member): the list's delegated click — the item hit pastes its entry. */
  onListClick(event: Event): void {
    const item = (event.target as HTMLElement | null)?.closest(".item");
    if (!item || !this.listEl) return;
    const entries = [...this.listEl.querySelectorAll(".item")];
    const entry = this.entries[entries.indexOf(item)];
    if (entry) this.emitPaste(entry, event);
  }

  pasteAll(): void {
    this.$emit("clipboard:paste-all");
  }

  clear(): void {
    this.entries = [];
    this.$emit("clipboard:clear");
  }

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    // Labels resolve once per template pass — re-assigning the observable
    // entries re-runs the bindings on a locale change (host data is kept).
    this.#unobserveLang = observeLang(() => {
      this.entries = [...this.entries];
    });
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#unobserveLang?.();
  }
}

export default DocenClipboardPane;
