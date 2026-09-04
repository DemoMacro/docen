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

/** One comment card's data — the host flattens documentExtras.comments into
 *  this shape (children runs joined to a single text). */
export interface CommentCard {
  id: number;
  author: string;
  initials: string;
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
  /* Card list — Word's comments pane: one card per comment, a rounded avatar
     with the author's initials, name + timestamp on one row, body under. */
  .list {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 6px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .empty {
    color: var(--docen-color-text-2, #616161);
    padding: 14px 8px;
    text-align: center;
  }
  .card {
    display: grid;
    grid-template-columns: 28px 1fr auto;
    grid-template-rows: auto auto;
    column-gap: 8px;
    padding: 8px;
    border-radius: 4px;
    cursor: pointer;
  }
  .card:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  /* The card whose anchored range the caret sits in — Word highlights it
     while the selection is inside the comment's text. */
  .card.active {
    background: var(--docen-color-subtle-selected, #e8f0fb);
  }
  .avatar {
    grid-row: 1 / 3;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    color: #fff;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 600;
  }
  .meta {
    display: flex;
    align-items: baseline;
    gap: 6px;
    min-width: 0;
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
  }
  .card:hover .actions {
    display: inline-flex;
  }
  .act {
    font-size: 11px;
    color: var(--docen-color-text-1, #242424);
  }
  .body {
    grid-column: 2 / 4;
    margin-top: 2px;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
  }
  /* Inline edit state — the card's body becomes a text area with
     Save / Cancel (Word edits in place, not through a dialog). */
  .edit {
    grid-column: 2 / 4;
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-top: 2px;
  }
  .edit fluent-textarea {
    width: 100%;
    box-sizing: border-box;
  }
  .edit .row {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }
`;

const template = html<DocenCommentsPane>`<div class="list" ${ref("listEl")}></div>`;

/** `<docen-comments-pane comments active-id>` — Word's comments sidebar: one
 *  card per comment (initials avatar, author, timestamp, body) and an inline
 *  edit state per card. `comments` is JSON `CommentCard[]` in the order the
 *  host computed (document order); `active-id` highlights the card whose
 *  anchored range the selection sits in. Interactions emit `comment:select
 *  {id}` (card click — the host scrolls to the range), `comment:update
 *  {id,text}` and `comment:delete {id}`. New comments compose in the floating
 *  box beside the selection (host-owned), not in this pane. */
@customElement({ name: "docen-comments-pane", template, styles })
class DocenCommentsPane extends FASTElement {
  /** JSON CommentCard[] — the document's comments, ordered by the host. */
  @attr comments?: string;
  /** The comment id whose range covers the selection ("" = none). */
  @attr({ attribute: "active-id" }) activeId?: string;

  @observable listEl?: HTMLElement;
  #unsubscribe?: () => void;

  commentsChanged(): void {
    this.#renderList();
  }

  activeIdChanged(): void {
    this.#highlightActive();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderList();
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  /** One card. The edit state swaps the body for a text area + Save/Cancel
   *  and is tracked per pane (#editingId), not per attr. */
  #renderCard(comment: CommentCard, frag: DocumentFragment): void {
    const card = document.createElement("div");
    card.className = "card";
    card.dataset.id = String(comment.id);
    card.addEventListener("click", (e) => {
      if ((e.target as HTMLElement).closest(".actions, .edit")) return;
      this.dispatchEvent(
        new CustomEvent("comment:select", {
          bubbles: true,
          composed: true,
          detail: { id: comment.id },
        }),
      );
    });

    const avatar = document.createElement("div");
    avatar.className = "avatar";
    avatar.textContent = comment.initials || comment.author.slice(0, 2).toUpperCase();
    card.append(avatar);

    const meta = document.createElement("div");
    meta.className = "meta";
    const author = document.createElement("span");
    author.className = "author";
    author.textContent = comment.author;
    const when = document.createElement("span");
    when.className = "when";
    try {
      when.textContent = new Date(comment.date).toLocaleString();
    } catch {
      when.textContent = comment.date;
    }
    meta.append(author, when);
    card.append(meta);

    const actions = document.createElement("div");
    actions.className = "actions";
    const edit = document.createElement("fluent-button");
    edit.setAttribute("appearance", "subtle");
    edit.setAttribute("size", "small");
    edit.className = "act";
    edit.textContent = t("comments.edit", this);
    edit.addEventListener("click", () => this.#beginEdit(comment, card));
    const del = document.createElement("fluent-button");
    del.setAttribute("appearance", "subtle");
    del.setAttribute("size", "small");
    del.className = "act";
    del.textContent = t("comments.delete", this);
    del.addEventListener("click", () => {
      this.dispatchEvent(
        new CustomEvent("comment:delete", {
          bubbles: true,
          composed: true,
          detail: { id: comment.id },
        }),
      );
    });
    actions.append(edit, del);
    card.append(actions);

    const body = document.createElement("div");
    body.className = "body";
    body.textContent = comment.text;
    card.append(body);

    frag.append(card);
  }

  #beginEdit(comment: CommentCard, card: HTMLElement): void {
    if (card.querySelector(".edit")) return;
    card.querySelector(".body")?.remove();
    const wrap = document.createElement("div");
    wrap.className = "edit";
    const area = document.createElement("fluent-textarea") as HTMLTextAreaElement & HTMLElement;
    // `block` drops Fluent's fixed 18rem inline-size — without it the inner
    // root box overflows the card (see the floating compose box).
    area.setAttribute("block", "");
    area.setAttribute("rows", "3");
    area.value = comment.text;
    const row = document.createElement("div");
    row.className = "row";
    const cancel = document.createElement("fluent-button");
    cancel.setAttribute("appearance", "neutral");
    cancel.textContent = t("comments.cancel", this);
    cancel.addEventListener("click", () => {
      this.#renderList();
    });
    const save = document.createElement("fluent-button");
    save.setAttribute("appearance", "accent");
    save.textContent = t("comments.save", this);
    save.addEventListener("click", () => {
      const text = (area.value ?? "").trim();
      if (text)
        this.dispatchEvent(
          new CustomEvent("comment:update", {
            bubbles: true,
            composed: true,
            detail: { id: comment.id, text },
          }),
        );
      this.#renderList();
    });
    area.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        save.click();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        this.#renderList();
      }
    });
    row.append(cancel, save);
    wrap.append(area, row);
    card.append(wrap);
    // Focus the shadow <textarea> — the host itself is not focusable.
    requestAnimationFrame(() => {
      const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
      input?.focus();
    });
  }

  #renderList(): void {
    const list = this.listEl;
    if (!list) return;
    list.replaceChildren();
    let cards: CommentCard[] = [];
    try {
      cards = this.comments ? (JSON.parse(this.comments) as CommentCard[]) : [];
    } catch {
      cards = [];
    }
    if (cards.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = t("comments.empty", this);
      list.append(empty);
      return;
    }
    const frag = document.createDocumentFragment();
    for (const card of cards) this.#renderCard(card, frag);
    list.append(frag);
    this.#highlightActive();
  }

  /** Paint the active card and bring it into the pane's viewport — the scroll
   *  stays inside .list (manual scrollTop), never the page behind the pane. */
  #highlightActive(): void {
    const list = this.listEl;
    if (!list) return;
    const id = this.activeId;
    for (const el of list.querySelectorAll<HTMLElement>(".card")) {
      const on = id != null && id !== "" && el.dataset.id === id;
      el.classList.toggle("active", on);
      if (on) list.scrollTop = el.offsetTop - list.clientHeight / 2 + el.clientHeight / 2;
    }
  }

  #applyI18n(): void {
    this.#renderList();
  }
}

export default DocenCommentsPane;
