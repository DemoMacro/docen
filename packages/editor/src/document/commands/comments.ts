import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { TextSelection } from "@tiptap/pm/state";

import { t } from "../../ui";

/** One inlinePassthrough comment atom → its marker kind and comment id;
 *  non-comment atoms yield null. Accepts both a JSON atom (type is the string
 *  name) and a PM node from doc.descendants (type is the NodeType — read its
 *  name). */
function commentMarkerOf(
  child: JSONContent | { type?: { name?: string }; attrs?: { data?: string } },
): { id: number; kind: "start" | "end" | "reference" } | null {
  const type = child.type as string | { name?: string } | undefined;
  const typeName = typeof type === "string" ? type : type?.name;
  if (typeName !== "inlinePassthrough") return null;
  try {
    const data = JSON.parse(String((child.attrs as { data?: string } | undefined)?.data)) as {
      commentRangeStart?: { id?: number };
      commentRangeEnd?: { id?: number };
      commentReference?: number;
    };
    if (typeof data.commentRangeStart?.id === "number")
      return { id: data.commentRangeStart.id, kind: "start" };
    if (typeof data.commentRangeEnd?.id === "number")
      return { id: data.commentRangeEnd.id, kind: "end" };
    if (typeof data.commentReference === "number")
      return { id: data.commentReference, kind: "reference" };
  } catch {
    // Malformed passthrough data — not a comment marker.
  }
  return null;
}

/** The word at `pos` — the non-whitespace run in the caret's text node, capped
 *  at 32 chars around the caret (CJK text has no spaces, so an uncapped run
 *  swallows the whole paragraph). Null when the caret touches no text. */
function wordRangeAt(
  doc: import("@tiptap/pm/model").Node,
  pos: number,
): { from: number; to: number } | null {
  const $pos = doc.resolve(pos);
  const node = $pos.nodeBefore?.isText
    ? $pos.nodeBefore
    : $pos.nodeAfter?.isText
      ? $pos.nodeAfter
      : null;
  if (!node) return null;
  const start = pos - ($pos.nodeBefore?.isText ? $pos.nodeBefore.nodeSize : 0);
  const chars = node.text ?? "";
  const at = pos - start;
  let from = at;
  while (from > 0 && !/\s/.test(chars.charAt(from - 1))) from--;
  let to = at;
  while (to < chars.length && !/\s/.test(chars.charAt(to))) to++;
  if (to - from > 32) {
    from = Math.max(from, at - 16);
    to = Math.min(to, at + 16);
  }
  return from < to ? { from: start + from, to: start + to } : null;
}

/** The comments commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface CommentsHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — the compose card anchors against the canvas, commits
   *  hand focus back, and card clicks scroll the anchored text into view. */
  bridge():
    | {
        focus(): void;
        scrollIntoView(pos: number): void;
        commentAnchorRect(
          from: number,
          to: number,
        ): { frame: HTMLElement; left: number; top: number } | null;
      }
    | undefined;
  /** The host element — the shadow-DOM root for the comments pane, the i18n
   *  language source, and the target of the comment:* events. */
  element(): HTMLElement;
  /** Reveal the comments pane (Review → Edit Comment edits inline on the
   *  card, Word's sidebar shape). */
  showTaskpane(id: "comments"): void;
}

/**
 * The comment domain, split out of the host element: anchoring a selection
 * with a commentRangeStart/End + Reference atom triple, the floating compose
 * card, the pane's card list sync (document-order cards), and the
 * select/update/delete/jump interactions (the markers are inlinePassthrough
 * atoms, the content lives in doc.attrs.documentExtras.comments).
 */
export class CommentsCommands {
  constructor(private readonly host: CommentsHost) {}

  /** The selection (or caret word) the pending compose will anchor. */
  #pendingCommentRange?: { from: number; to: number };

  /** The floating compose box over the canvas (null once dismissed). */
  #commentCompose?: HTMLElement;

  /** The comment whose range covers the selection (a marker pair bracketing
   *  it in document order — markers may sit in earlier paragraphs), lowest id
   *  first; null when the selection touches no comment. */
  activeCommentId(): number | null {
    const editor = this.host.editor();
    if (!editor) return null;
    const { from, to } = editor.state.selection;
    const opened = new Map<number, number>();
    const covering = new Set<number>();
    editor.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (!marker) return;
      if (marker.kind === "start") opened.set(marker.id, pos);
      else if (marker.kind === "end") {
        const start = opened.get(marker.id);
        if (start != null && start < to && pos > from) covering.add(marker.id);
      }
    });
    return covering.size > 0 ? Math.min(...covering) : null;
  }

  /** Review → Edit Comment: open the comments pane — editing happens inline
   *  on the card (Word edits comments in the sidebar, not a dialog). */
  editComment(): void {
    this.host.showTaskpane("comments");
  }

  /** Review → Delete: remove the covering comment's marker/reference atoms
   *  (descending positions keep the earlier offsets valid) and its
   *  documentExtras entry. */
  deleteComment(): void {
    const id = this.activeCommentId();
    if (id == null) return;
    this.#deleteCommentById(id);
  }

  /** Remove a comment's marker/reference atoms (descending positions keep the
   *  earlier offsets valid) and its documentExtras entry — shared by the
   *  ribbon's Delete Comment and the pane's per-card delete. */
  #deleteCommentById(id: number): void {
    const editor = this.host.editor();
    if (!editor) return;
    const atoms: { pos: number; size: number }[] = [];
    editor.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (marker && marker.id === id) atoms.push({ pos, size: child.nodeSize });
    });
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Record<string, unknown>[] };
    };
    const tr = editor.state.tr;
    for (const { pos, size } of atoms.sort((a, b) => b.pos - a.pos)) tr.delete(pos, pos + size);
    tr.setDocAttribute("documentExtras", {
      ...docAttrs.documentExtras,
      comments: (docAttrs.documentExtras?.comments ?? []).filter((c) => Number(c.id) !== id),
    });
    editor.view.dispatch(tr);
  }

  /** Document Inspector → Remove All: every comment's marker/reference atoms
   *  and the whole documentExtras list go in ONE transaction (one undo step) —
   *  deleting them one by one would also re-render the pane N times. */
  deleteAllComments(): void {
    const editor = this.host.editor();
    if (!editor) return;
    const atoms: { pos: number; size: number }[] = [];
    editor.state.doc.descendants((child, pos) => {
      if (commentMarkerOf(child)) atoms.push({ pos, size: child.nodeSize });
    });
    if (atoms.length === 0) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Record<string, unknown>[] };
    };
    const tr = editor.state.tr;
    for (const { pos, size } of atoms.sort((a, b) => b.pos - a.pos)) tr.delete(pos, pos + size);
    tr.setDocAttribute("documentExtras", {
      ...docAttrs.documentExtras,
      comments: [],
    });
    editor.view.dispatch(tr);
  }

  /** Review → Previous/Next Comment: select the range of the comment before
   *  or after the selection (document order); no further comment in that
   *  direction is a no-op. */
  jumpComment(direction: "previous" | "next"): void {
    const editor = this.host.editor();
    if (!editor) return;
    const ranges: { from: number; to: number }[] = [];
    const openStarts = new Map<number, number>();
    editor.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (!marker) return;
      if (marker.kind === "start") openStarts.set(marker.id, pos + child.nodeSize);
      else if (marker.kind === "end") {
        const start = openStarts.get(marker.id);
        if (start != null) ranges.push({ from: start, to: pos });
      }
    });
    ranges.sort((a, b) => a.from - b.from);
    const { from } = editor.state.selection;
    const target =
      direction === "next"
        ? ranges.find((r) => r.from > from)
        : ranges.findLast((r) => r.from < from);
    if (!target) return;
    editor.view.dispatch(
      editor.state.tr.setSelection(
        new TextSelection(
          editor.state.doc.resolve(target.from),
          editor.state.doc.resolve(target.to),
        ),
      ),
    );
  }

  /** Review → New Comment: open the floating compose box beside the
   *  selection (Word's Simple-Markup reply card — it hangs in the margin at
   *  the anchored line and scrolls with its page). The text arrives via the
   *  `comment:create` event (#onCommentCreate commits it). Without a
   *  selection the word at the caret anchors the comment (Word for the web's
   *  behavior); a caret on whitespace is a no-op. */
  insertComment(): void {
    const editor = this.host.editor();
    if (!editor) return;
    // Spread would miss from/to — they're prototype getters on Selection.
    const { from, to } = editor.state.selection;
    if (from === to) {
      const word = wordRangeAt(editor.state.doc, from);
      if (!word) return;
      this.#pendingCommentRange = word;
    } else {
      this.#pendingCommentRange = { from, to };
    }
    this.#openCommentCompose();
  }

  /** Mount the floating compose card on the anchored page frame — Fluent
   *  components (text-area, buttons) over the design-token palette, no
   *  hand-rolled colors. */
  #openCommentCompose(): void {
    this.#closeCommentCompose();
    const range = this.#pendingCommentRange;
    if (!range) return;
    const anchor = this.host.bridge()?.commentAnchorRect(range.from, range.to);
    if (!anchor) {
      // Unmappable selection (stale map) — drop the pending compose.
      this.#pendingCommentRange = undefined;
      return;
    }
    const card = document.createElement("div");
    card.setAttribute("data-docen-overlay", "");
    Object.assign(card.style, {
      position: "absolute",
      zIndex: "6",
      width: "220px",
      boxSizing: "border-box",
      padding: "10px",
      display: "flex",
      flexDirection: "column",
      gap: "8px",
      background: "var(--docen-color-bg, #ffffff)",
      border: "1px solid var(--docen-color-divider, #e2e2e2)",
      borderRadius: "var(--borderRadiusLarge, 8px)",
      boxShadow: "var(--shadow4, 0 4px 8px rgba(0,0,0,.14))",
      fontFamily: "inherit",
      fontSize: "var(--docen-font-size-ribbon, 12px)",
    } satisfies Partial<CSSStyleDeclaration>);
    const area = document.createElement("fluent-textarea") as HTMLTextAreaElement & HTMLElement;
    // `block` drops Fluent's fixed 18rem inline-size — without it the inner
    // root box overflows the 220px card.
    area.setAttribute("block", "");
    area.setAttribute("resize", "vertical");
    area.setAttribute("rows", "3");
    area.setAttribute("placeholder", t("comments.placeholder", this.host.element()));
    const row = document.createElement("div");
    Object.assign(row.style, {
      display: "flex",
      gap: "6px",
      justifyContent: "flex-end",
    } satisfies Partial<CSSStyleDeclaration>);
    const cancel = document.createElement("fluent-button");
    cancel.setAttribute("appearance", "neutral");
    cancel.textContent = t("comments.cancel", this.host.element());
    const post = document.createElement("fluent-button");
    post.setAttribute("appearance", "accent");
    post.textContent = t("comments.post", this.host.element());
    cancel.addEventListener("click", () =>
      this.host
        .element()
        .dispatchEvent(new CustomEvent("comment:cancel", { bubbles: true, composed: true })),
    );
    const postIt = (): void => {
      const text = (area.value ?? "").trim();
      if (!text) return;
      this.host
        .element()
        .dispatchEvent(
          new CustomEvent("comment:create", { bubbles: true, composed: true, detail: { text } }),
        );
    };
    post.addEventListener("click", postIt);
    // Ctrl+Enter posts (the sidebar edit's shortcut); Escape cancels.
    area.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        postIt();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        this.host
          .element()
          .dispatchEvent(new CustomEvent("comment:cancel", { bubbles: true, composed: true }));
      }
    });
    row.append(cancel, post);
    card.append(area, row);
    // Word hangs the card in the margin just past the anchored line's end,
    // clamped into the page frame when the line runs to the right edge.
    const frameW = anchor.frame.clientWidth;
    const CARD_W = 220;
    const left = Math.min(anchor.left + 12, Math.max(frameW - CARD_W - 8, 0));
    Object.assign(card.style, {
      left: `${left}px`,
      top: `${anchor.top}px`,
    } satisfies Partial<CSSStyleDeclaration>);
    anchor.frame.append(card);
    this.#commentCompose = card;
    // The host is not focusable — focus lands on the shadow <textarea>, or
    // every keystroke goes to the canvas bridge and into the document.
    requestAnimationFrame(() => {
      const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
      input?.focus();
    });
  }

  /** Tear down the floating compose card (Post, Cancel, or Escape). */
  #closeCommentCompose(): void {
    this.#commentCompose?.remove();
    this.#commentCompose = undefined;
  }

  /** comment:create → commit the pending comment: anchor the stored range the
   *  way Word does — a commentRangeStart/commentRangeEnd passthrough pair
   *  around it with a commentReference after — and append the structured
   *  content to doc.attrs.documentExtras.comments (word/comments.xml on
   *  export, the round-trip channel the parse side already fills). */
  readonly onCommentCreate = (event: CustomEvent<{ text?: string }>): void => {
    const editor = this.host.editor();
    const text = event.detail?.text?.trim();
    const range = this.#pendingCommentRange;
    this.#pendingCommentRange = undefined;
    this.#closeCommentCompose();
    if (!editor || !text || !range) return;
    // Word returns the caret to the body once the comment is committed.
    this.host.bridge()?.focus();
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Record<string, unknown>[] };
    };
    const comments = docAttrs.documentExtras?.comments ?? [];
    const id = comments.reduce((max, c) => Math.max(max, Number(c.id ?? 0)), -1) + 1;
    const seed = (data: object): JSONContent =>
      ({
        type: "inlinePassthrough",
        attrs: { data: JSON.stringify(data) },
      }) as JSONContent;
    const start = editor.schema.nodeFromJSON(seed({ commentRangeStart: { id } }));
    const end = editor.schema.nodeFromJSON(seed({ commentRangeEnd: { id } }));
    const ref = editor.schema.nodeFromJSON(seed({ commentReference: id }));
    editor.view.dispatch(
      editor.state.tr
        .insert(range.from, start)
        .insert(range.to + 1, end)
        .insert(range.to + 2, ref)
        .setDocAttribute("documentExtras", {
          ...docAttrs.documentExtras,
          comments: [
            ...comments,
            {
              id,
              author: "Docen User",
              initials: "DU",
              date: new Date().toISOString(),
              children: [{ text }],
            },
          ],
        }),
    );
  };

  /** comment:cancel → drop the pending compose (Word's Cancel). */
  readonly onCommentCancel = (): void => {
    this.#pendingCommentRange = undefined;
    this.#closeCommentCompose();
  };

  /** comment:select → select and scroll to the comment's range (Word scrolls
   *  the anchored text into view when its card is clicked). */
  readonly onCommentSelect = (event: CustomEvent<{ id?: number }>): void => {
    const editor = this.host.editor();
    const id = event.detail?.id;
    if (!editor || id == null) return;
    const range = this.#commentRangeOf(id);
    if (!range) return;
    editor.view.dispatch(
      editor.state.tr.setSelection(
        new TextSelection(editor.state.doc.resolve(range.from), editor.state.doc.resolve(range.to)),
      ),
    );
    this.host.bridge()?.scrollIntoView(range.from);
  };

  /** comment:update → rewrite the comment's text (the sidebar's inline edit,
   *  replacing the old prompt-based Edit Comment for pane interactions). */
  readonly onCommentUpdate = (event: CustomEvent<{ id?: number; text?: string }>): void => {
    const editor = this.host.editor();
    const id = event.detail?.id;
    const text = event.detail?.text?.trim();
    if (!editor || id == null || !text) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as {
      documentExtras?: { comments?: Array<{ id?: number; children?: Array<{ text?: string }> }> };
    };
    const comments = docAttrs.documentExtras?.comments ?? [];
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", {
        ...docAttrs.documentExtras,
        comments: comments.map((c) => (Number(c.id) === id ? { ...c, children: [{ text }] } : c)),
      }),
    );
  };

  /** comment:delete → remove the comment's marker/reference atoms and its
   *  documentExtras entry (same teardown as the ribbon's Delete Comment). */
  readonly onCommentDelete = (event: CustomEvent<{ id?: number }>): void => {
    const id = event.detail?.id;
    if (id != null) this.#deleteCommentById(id);
  };

  /** The comments pane element (inside its task-pane), when mounted. */
  #commentsPaneEl(): (HTMLElement & { comments?: string; activeId?: string }) | null {
    return this.host.element().shadowRoot?.querySelector("docen-comments-pane") ?? null;
  }

  /** The document order range a comment id covers (start marker through end
   *  marker), or null when its markers are gone. */
  #commentRangeOf(id: number): { from: number; to: number } | null {
    const editor = this.host.editor();
    if (!editor) return null;
    let from: number | null = null;
    let to: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (!marker || marker.id !== id) return;
      if (marker.kind === "start") from = pos + child.nodeSize;
      else if (marker.kind === "end") to = pos;
    });
    return from != null && to != null ? { from, to } : null;
  }

  /** Sync the comments pane's card list from the doc's documentExtras after
   *  every transaction (the pane is a pure view of the model). Cards order by
   *  their anchored range's document position — Word's sidebar follows the
   *  text, not the round-trip append order the extras array stores. */
  syncCommentsPane(): void {
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    const editor = this.host.editor();
    const docAttrs = (editor?.state.doc.attrs ?? {}) as {
      documentExtras?: {
        comments?: Array<{
          id?: number;
          author?: string;
          initials?: string;
          date?: string;
          children?: Array<{ text?: string }>;
        }>;
      };
    };
    const startPos = new Map<number, number>();
    editor?.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (marker?.kind === "start") startPos.set(marker.id, pos);
    });
    const cards = (docAttrs.documentExtras?.comments ?? [])
      .map((c) => ({
        id: Number(c.id ?? 0),
        author: c.author ?? "",
        initials: c.initials ?? "",
        date: c.date ?? "",
        text: (c.children ?? []).map((r) => r.text ?? "").join(""),
        pos: startPos.get(Number(c.id ?? 0)),
      }))
      .sort((a, b) => (a.pos ?? Number.MAX_SAFE_INTEGER) - (b.pos ?? Number.MAX_SAFE_INTEGER))
      .map(({ pos: _pos, ...card }) => card);
    pane.comments = JSON.stringify(cards);
  }

  /** selectionUpdate → highlight the comments-pane card whose anchored range
   *  covers the selection (Word paints the anchored card while the caret sits
   *  in its text). */
  readonly syncActiveCommentCard = (): void => {
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    const id = this.activeCommentId();
    pane.setAttribute("active-id", id == null ? "" : String(id));
  };
}
