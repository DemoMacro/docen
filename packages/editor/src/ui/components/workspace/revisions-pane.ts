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

/** One revision card's data — the host flattens collectRevisions() into this
 *  shape (index = the entry the accept/reject events refer to). */
export interface RevisionCard {
  index: number;
  type: "insertion" | "deletion";
  author: string;
  date: string;
  text: string;
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
  /* Word's reviewing pane: one row per revision — a type tag, author, time,
     and the tracked text — with accept/reject on hover. */
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
  .card {
    display: grid;
    grid-template-columns: 1fr auto;
    grid-template-rows: auto auto;
    column-gap: 8px;
    padding: 6px 8px;
    border-radius: 4px;
    cursor: pointer;
  }
  .card:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .card.active {
    background: var(--docen-color-subtle-selected, #e8f0fb);
  }
  .meta {
    display: flex;
    align-items: baseline;
    gap: 6px;
    min-width: 0;
  }
  .kind {
    font-size: 10px;
    line-height: 16px;
    padding: 0 6px;
    border-radius: 2px;
    white-space: nowrap;
  }
  .kind.insertion {
    background: var(--docen-color-status-success-background, #e6f2ea);
    color: var(--docen-color-status-success, #0e700e);
  }
  .kind.deletion {
    background: var(--docen-color-status-warning-background, #fdf6e6);
    color: var(--docen-color-status-warning, #8a5b00);
  }
  .author {
    font-weight: 600;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .when {
    color: var(--docen-color-text-2, #616161);
    white-space: nowrap;
    font-size: 11px;
  }
  .actions {
    display: none;
    gap: 2px;
    grid-row: 1;
    grid-column: 2;
  }
  .card:hover .actions {
    display: inline-flex;
  }
  .act {
    font-size: 11px;
    color: var(--docen-color-text-1, #242424);
  }
  .text {
    grid-column: 1 / 3;
    margin-top: 2px;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    color: var(--docen-color-text-1, #242424);
  }
  .text.empty-text {
    color: var(--docen-color-text-2, #616161);
    font-style: italic;
  }
`;

const template = html<DocenRevisionsPane>`<div class="list" ${ref("listEl")}></div>`;

/** `<docen-revisions-pane revisions active-index>` — Word's reviewing pane:
 *  one card per tracked change (insertion/deletion tag, author, timestamp,
 *  text) with hover accept/reject. `revisions` is JSON `RevisionCard[]` in
 *  document order; `active-index` highlights the card whose range covers the
 *  selection. Interactions emit `revision:select {index}` (card click — the
 *  host moves the selection), `revision:accept {index}` and
 *  `revision:reject {index}`. */
@customElement({ name: "docen-revisions-pane", template, styles })
class DocenRevisionsPane extends FASTElement {
  /** JSON RevisionCard[] — the document's revisions, in document order. */
  @attr revisions?: string;
  /** The revision index whose range covers the selection ("" = none). */
  @attr({ attribute: "active-index" }) activeIndex?: string;

  @observable listEl?: HTMLElement;
  #unsubscribe?: () => void;

  revisionsChanged(): void {
    this.#renderList();
  }

  activeIndexChanged(): void {
    this.#highlightActive();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderList();
    this.#unsubscribe = observeLang(() => this.#renderList());
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    this.#unsubscribe = undefined;
    super.disconnectedCallback();
  }

  #renderCard(card: RevisionCard, frag: DocumentFragment): void {
    const el = document.createElement("div");
    el.className = "card";
    el.dataset.index = String(card.index);
    el.addEventListener("click", (e) => {
      if ((e.target as HTMLElement).closest(".actions")) return;
      this.dispatchEvent(
        new CustomEvent("revision:select", {
          bubbles: true,
          composed: true,
          detail: { index: card.index },
        }),
      );
    });

    const meta = document.createElement("div");
    meta.className = "meta";
    const kind = document.createElement("span");
    kind.className = `kind ${card.type}`;
    kind.textContent = t(`revisions.${card.type}`, this);
    const author = document.createElement("span");
    author.className = "author";
    author.textContent = card.author;
    const when = document.createElement("span");
    when.className = "when";
    try {
      when.textContent = new Date(card.date).toLocaleString();
    } catch {
      when.textContent = card.date;
    }
    meta.append(kind, author, when);
    el.append(meta);

    const actions = document.createElement("div");
    actions.className = "actions";
    const accept = document.createElement("fluent-button");
    accept.setAttribute("appearance", "subtle");
    accept.setAttribute("size", "small");
    accept.className = "act";
    accept.textContent = t("revisions.accept", this);
    accept.addEventListener("click", () => {
      this.dispatchEvent(
        new CustomEvent("revision:accept", {
          bubbles: true,
          composed: true,
          detail: { index: card.index },
        }),
      );
    });
    const reject = document.createElement("fluent-button");
    reject.setAttribute("appearance", "subtle");
    reject.setAttribute("size", "small");
    reject.className = "act";
    reject.textContent = t("revisions.reject", this);
    reject.addEventListener("click", () => {
      this.dispatchEvent(
        new CustomEvent("revision:reject", {
          bubbles: true,
          composed: true,
          detail: { index: card.index },
        }),
      );
    });
    actions.append(accept, reject);
    el.append(actions);

    const text = document.createElement("div");
    const value = card.text;
    text.className = value ? "text" : "text empty-text";
    text.textContent = value || t("revisions.emptyText", this);
    el.append(text);

    frag.append(el);
  }

  #renderList(): void {
    const list = this.listEl;
    if (!list) return;
    list.replaceChildren();
    let cards: RevisionCard[] = [];
    try {
      cards = this.revisions ? (JSON.parse(this.revisions) as RevisionCard[]) : [];
    } catch {
      cards = [];
    }
    if (cards.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = t("revisions.empty", this);
      list.append(empty);
      return;
    }
    const frag = document.createDocumentFragment();
    for (const card of cards) this.#renderCard(card, frag);
    list.append(frag);
    this.#highlightActive();
  }

  /** Paint the active card and bring it into the pane's viewport — the scroll
   *  stays inside .list, never the page behind the pane. */
  #highlightActive(): void {
    const list = this.listEl;
    if (!list) return;
    const index = this.activeIndex;
    for (const el of list.querySelectorAll<HTMLElement>(".card")) {
      const on = index != null && index !== "" && el.dataset.index === index;
      el.classList.toggle("active", on);
      if (on) list.scrollTop = el.offsetTop - list.clientHeight / 2 + el.clientHeight / 2;
    }
  }
}

export default DocenRevisionsPane;
