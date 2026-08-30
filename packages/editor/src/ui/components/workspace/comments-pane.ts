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
    border: none;
    background: transparent;
    color: var(--docen-color-text-1, #242424);
    font-size: 11px;
    padding: 2px 6px;
    border-radius: 3px;
    cursor: pointer;
  }
  .act:hover {
    background: var(--docen-color-stroke-1, #e0e0e0);
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
  .edit fluent-text-area {
    width: 100%;
    box-sizing: border-box;
  }
  .edit .row {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }
  /* Compose box — shown while a New Comment is pending (Word opens the
     sidebar's compose card with Post / Cancel; Ctrl+Enter posts). */
  .compose {
    flex: 0 0 auto;
    padding: 8px;
    border-top: 1px solid var(--docen-color-divider, #e2e2e2);
    display: none;
    flex-direction: column;
    gap: 6px;
  }
  :host([draft]) .compose {
    display: flex;
  }
  .compose fluent-text-area {
    width: 100%;
    box-sizing: border-box;
  }
  .compose .row {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }
`;

const template = html<DocenCommentsPane>`
  <div class="list" ${ref("listEl")}></div>
  <div class="compose" ${ref("composeEl")}>
    <fluent-text-area ${ref("composeArea")} resize="vertical" rows="3"></fluent-text-area>
    <div class="row">
      <fluent-button
        appearance="neutral"
        ${ref("composeCancel")}
        data-i18n="comments.cancel"
      ></fluent-button>
      <fluent-button appearance="accent" ${ref("composePost")} data-i18n="comments.post">
      </fluent-button>
    </div>
  </div>
`;

/** `<docen-comments-pane comments draft>` — Word's comments sidebar: one card
 *  per comment (initials avatar, author, timestamp, body), an inline edit
 *  state per card, and a compose box at the bottom while `draft` is set.
 *  `comments` is JSON `CommentCard[]`; interactions emit `comment:select
 *  {id}` (card click — the host scrolls to the range), `comment:create
 *  {text}` (Post / Ctrl+Enter), `comment:cancel`, `comment:update {id,text}`
 *  and `comment:delete {id}`. */
@customElement({ name: "docen-comments-pane", template, styles })
class DocenCommentsPane extends FASTElement {
  /** JSON CommentCard[] — the document's comments, flattened by the host. */
  @attr comments?: string;
  /** Show the compose box (a New Comment is pending its text). */
  @attr({ mode: "boolean" }) draft?: boolean;

  @observable listEl?: HTMLElement;
  @observable composeEl?: HTMLElement;
  @observable composeArea?: HTMLTextAreaElement & HTMLElement;
  @observable composePost?: HTMLElement;
  @observable composeCancel?: HTMLElement;
  #unsubscribe?: () => void;

  commentsChanged(): void {
    this.#renderList();
  }

  draftChanged(): void {
    if (this.draft) {
      // The compose box opens focused, like Word's new-comment card.
      requestAnimationFrame(() => this.composeArea?.focus());
    }
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderList();
    this.#applyI18n();
    this.composePost?.addEventListener("click", () => this.#post());
    this.composeCancel?.addEventListener("click", () =>
      this.dispatchEvent(new CustomEvent("comment:cancel", { bubbles: true, composed: true })),
    );
    // Ctrl+Enter posts (Word's shortcut); a bare Enter keeps the newline so
    // comments can be multiline.
    this.composeArea?.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        this.#post();
      }
    });
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  #post(): void {
    const text = (this.composeArea?.value ?? "").trim();
    if (!text) return;
    this.composeArea!.value = "";
    this.dispatchEvent(
      new CustomEvent("comment:create", { bubbles: true, composed: true, detail: { text } }),
    );
  }

  /** One card. The edit state swaps the body for a text area + Save/Cancel
   *  and is tracked per pane (#editingId), not per attr. */
  #renderCard(comment: CommentCard, frag: DocumentFragment): void {
    const card = document.createElement("div");
    card.className = "card";
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
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "act";
    edit.textContent = t("comments.edit", this);
    edit.addEventListener("click", () => this.#beginEdit(comment, card));
    const del = document.createElement("button");
    del.type = "button";
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
    const area = document.createElement("fluent-text-area") as HTMLTextAreaElement & HTMLElement;
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
    requestAnimationFrame(() => area.focus());
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
  }

  #applyI18n(): void {
    this.composeArea?.setAttribute("placeholder", t("comments.placeholder", this));
    const root = this.shadowRoot;
    if (!root) return;
    root
      .querySelector('[data-i18n="comments.cancel"]')
      ?.replaceChildren(t("comments.cancel", this));
    root.querySelector('[data-i18n="comments.post"]')?.replaceChildren(t("comments.post", this));
    this.#renderList();
  }
}

export default DocenCommentsPane;
