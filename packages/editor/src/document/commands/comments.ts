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

/** The comment threads' model view — doc.attrs.documentExtras carries the
 *  round-trip channels: comments (word/comments.xml entries, replies included)
 *  and commentsExtended (w15:commentEx — the paraId reply links + resolved
 *  flags). Shapes mirror office-open's CommentOptions/CommentExtendedOptions. */
interface CommentExtras {
  comments?: Array<{
    id?: number;
    author?: string;
    initials?: string;
    date?: string;
    children?: unknown[];
  }>;
  commentsExtended?: Array<{
    paraId?: string;
    paraIdParent?: string;
    done?: boolean;
  }>;
}

/** The comments payload as stored in documentExtras (one w:comment). The
 *  first child is the paragraph that carries the thread-linking w14:paraId. */
type CommentEntry = NonNullable<CommentExtras["comments"]>[number];

/** A fresh w14:paraId (8-digit hex) not colliding with the given entries. */
function newParaId(taken: Iterable<string | undefined>): string {
  const used = new Set(taken);
  let id: string;
  do {
    id = Array.from({ length: 8 }, () =>
      "0123456789ABCDEF".charAt(Math.floor(Math.random() * 16)),
    ).join("");
  } while (used.has(id));
  return id;
}

/** The w14:paraId of a comment's first paragraph (null for legacy entries
 *  written before threads existed — they mint one on first thread operation). */
function paraIdOf(comment: CommentEntry): string | null {
  const first = comment.children?.[0];
  if (first && typeof first === "object") {
    const paraId = (first as { paraId?: unknown }).paraId;
    if (typeof paraId === "string") return paraId;
  }
  return null;
}

/** A copy of the comment whose first paragraph carries paraId. */
function withParaId(comment: CommentEntry, paraId: string): CommentEntry {
  const [first, ...rest] = comment.children ?? [];
  const head =
    first === undefined
      ? { children: [], paraId }
      : typeof first === "string"
        ? { children: [first], paraId }
        : { ...first, paraId };
  return { ...comment, children: [head, ...rest] };
}

/** A comment's plain text — children may be raw strings, run shapes ({text})
 *  or paragraph shapes ({children: [...]}) depending on who wrote the entry. */
function commentText(comment: CommentEntry): string {
  const walk = (kids: unknown[]): string =>
    (kids ?? [])
      .map((kid) => {
        if (typeof kid === "string") return kid;
        const obj = kid as { text?: unknown; children?: unknown[] };
        if (typeof obj.text === "string") return obj.text;
        if (Array.isArray(obj.children)) return walk(obj.children);
        return "";
      })
      .join("");
  return walk(comment.children ?? []);
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
   *  ribbon's Delete Comment and the pane's per-card delete. Deleting a thread
   *  root takes its replies with it (Word's behavior); deleting a reply keeps
   *  the rest of the thread. */
  #deleteCommentById(id: number): void {
    const editor = this.host.editor();
    if (!editor) return;
    const atoms: { pos: number; size: number }[] = [];
    editor.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (marker && marker.id === id) atoms.push({ pos, size: child.nodeSize });
    });
    const docAttrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const extras = docAttrs.documentExtras ?? {};
    const comments = extras.comments ?? [];
    const exs = extras.commentsExtended ?? [];
    const root = comments.find((c) => Number(c.id) === id);
    const threadParaIds = root ? this.#threadParaIds(root, comments, exs) : new Set<string>();
    const dropIds = new Set<number>([id]);
    for (const c of comments) {
      const pid = paraIdOf(c);
      if (pid && threadParaIds.has(pid)) dropIds.add(Number(c.id));
    }
    const tr = editor.state.tr;
    for (const { pos, size } of atoms.sort((a, b) => b.pos - a.pos)) tr.delete(pos, pos + size);
    tr.setDocAttribute("documentExtras", {
      ...extras,
      comments: comments.filter((c) => !dropIds.has(Number(c.id))),
      commentsExtended: exs.filter((e) => !e.paraId || !threadParaIds.has(e.paraId)),
    });
    editor.view.dispatch(tr);
  }

  /** Every paraId in the thread rooted at `root` (the root's own plus all
   *  descendants through commentsExtended's paraIdParent links). The walk is
   *  bounded per pass so a malformed cyclic chain can't spin forever. */
  #threadParaIds(
    root: CommentEntry,
    comments: CommentEntry[],
    exs: NonNullable<CommentExtras["commentsExtended"]>,
  ): Set<string> {
    const paraIds = new Set<string>();
    const rootParaId = paraIdOf(root);
    if (rootParaId) paraIds.add(rootParaId);
    let grown = true;
    while (grown) {
      grown = false;
      for (const c of comments) {
        const pid = paraIdOf(c);
        if (!pid || paraIds.has(pid)) continue;
        const parentParaId = exs.find((e) => e.paraId === pid)?.paraIdParent;
        if (parentParaId && paraIds.has(parentParaId)) {
          paraIds.add(pid);
          grown = true;
        }
      }
    }
    return paraIds;
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
    const docAttrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const tr = editor.state.tr;
    for (const { pos, size } of atoms.sort((a, b) => b.pos - a.pos)) tr.delete(pos, pos + size);
    tr.setDocAttribute("documentExtras", {
      ...docAttrs.documentExtras,
      comments: [],
      commentsExtended: [],
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
      documentExtras?: CommentExtras;
    };
    const extras = docAttrs.documentExtras ?? {};
    const comments = extras.comments ?? [];
    const exs = extras.commentsExtended ?? [];
    const id = comments.reduce((max, c) => Math.max(max, Number(c.id ?? 0)), -1) + 1;
    // The first paragraph's w14:paraId links the commentsExtended entry (the
    // Word 2013+ channel Word reads reply threading and resolved state from).
    const paraId = newParaId(exs.map((e) => e.paraId));
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
          ...extras,
          comments: [
            ...comments,
            {
              id,
              author: "Docen User",
              initials: "DU",
              date: new Date().toISOString(),
              children: [{ children: [text], paraId }],
            },
          ],
          commentsExtended: [...exs, { paraId, done: false }],
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

  /** comment:update → rewrite the comment's text (the sidebar's inline edit).
   *  The thread-linking paraId on the first paragraph survives the rewrite. */
  readonly onCommentUpdate = (event: CustomEvent<{ id?: number; text?: string }>): void => {
    const editor = this.host.editor();
    const id = event.detail?.id;
    const text = event.detail?.text?.trim();
    if (!editor || id == null || !text) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const comments = docAttrs.documentExtras?.comments ?? [];
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", {
        ...docAttrs.documentExtras,
        comments: comments.map((c) => {
          if (Number(c.id) !== id) return c;
          const paraId = paraIdOf(c);
          return { ...c, children: [paraId ? { children: [text], paraId } : text] };
        }),
      }),
    );
  };

  /** comment:reply → append a reply to the thread (Word's reply is its own
   *  w:comment with no body anchors, linked through commentsExtended's
   *  paraIdParent). Legacy parents without a paraId mint one so the reply can
   *  link to them — one transaction, one undo step. */
  readonly onCommentReply = (event: CustomEvent<{ parentId?: number; text?: string }>): void => {
    const editor = this.host.editor();
    const parentId = event.detail?.parentId;
    const text = event.detail?.text?.trim();
    if (!editor || parentId == null || !text) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const extras = docAttrs.documentExtras ?? {};
    let comments = extras.comments ?? [];
    let exs = extras.commentsExtended ?? [];
    const parent = comments.find((c) => Number(c.id) === parentId);
    if (!parent) return;
    let parentParaId = paraIdOf(parent);
    if (!parentParaId) {
      parentParaId = newParaId(exs.map((e) => e.paraId));
      comments = comments.map((c) =>
        Number(c.id) === parentId ? withParaId(c, parentParaId as string) : c,
      );
      exs = [...exs, { paraId: parentParaId, done: false }];
    }
    const id = comments.reduce((max, c) => Math.max(max, Number(c.id ?? 0)), -1) + 1;
    const paraId = newParaId(exs.map((e) => e.paraId));
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", {
        ...extras,
        comments: [
          ...comments,
          {
            id,
            author: "Docen User",
            initials: "DU",
            date: new Date().toISOString(),
            children: [{ children: [text], paraId }],
          },
        ],
        commentsExtended: [...exs, { paraId, paraIdParent: parentParaId, done: false }],
      }),
    );
  };

  /** comment:resolve → Word resolves the whole conversation: every
   *  commentsExtended entry in the thread gets the w15:done flag (or loses it
   *  on reopen). */
  readonly onCommentResolve = (event: CustomEvent<{ id?: number; done?: boolean }>): void => {
    const editor = this.host.editor();
    const id = event.detail?.id;
    const done = event.detail?.done === true;
    if (!editor || id == null) return;
    const docAttrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const extras = docAttrs.documentExtras ?? {};
    let comments = extras.comments ?? [];
    let exs = extras.commentsExtended ?? [];
    const root = comments.find((c) => Number(c.id) === id);
    if (!root) return;
    // Legacy roots mint their paraId so thread membership is well-defined.
    let rootParaId = paraIdOf(root);
    if (!rootParaId) {
      rootParaId = newParaId(exs.map((e) => e.paraId));
      comments = comments.map((c) =>
        Number(c.id) === id ? withParaId(c, rootParaId as string) : c,
      );
      exs = [...exs, { paraId: rootParaId, done: false }];
    }
    const threadParaIds = this.#threadParaIds(root, comments, exs);
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", {
        ...extras,
        comments,
        commentsExtended: exs.map((e) =>
          e.paraId && threadParaIds.has(e.paraId) ? { ...e, done } : e,
        ),
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
   *  every transaction (the pane is a pure view of the model). Threads order
   *  by their anchored range's document position — Word's sidebar follows the
   *  text, not the round-trip append order the extras array stores; replies
   *  hang under their thread root through the commentsExtended paraIdParent
   *  chain (multi-level links flatten under the root). */
  syncCommentsPane(): void {
    const pane = this.#commentsPaneEl();
    if (!pane) return;
    const editor = this.host.editor();
    const docAttrs = (editor?.state.doc.attrs ?? {}) as { documentExtras?: CommentExtras };
    const comments = docAttrs.documentExtras?.comments ?? [];
    const exs = docAttrs.documentExtras?.commentsExtended ?? [];
    const exByParaId = new Map(exs.filter((e) => e.paraId).map((e) => [e.paraId as string, e]));
    const startPos = new Map<number, number>();
    editor?.state.doc.descendants((child, pos) => {
      const marker = commentMarkerOf(child);
      if (marker?.kind === "start") startPos.set(marker.id, pos);
    });
    interface ThreadCard {
      id: number;
      author: string;
      initials: string;
      date: string;
      text: string;
      resolved: boolean;
      replies: ThreadCard[];
    }
    const cardOf = new Map<number, ThreadCard>();
    const makeCard = (c: CommentEntry): ThreadCard => {
      const paraId = paraIdOf(c);
      const card: ThreadCard = {
        id: Number(c.id ?? 0),
        author: c.author ?? "",
        initials: c.initials ?? "",
        date: c.date ?? "",
        text: commentText(c),
        resolved: paraId != null && exByParaId.get(paraId)?.done === true,
        replies: [],
      };
      cardOf.set(card.id, card);
      return card;
    };
    const ordered = comments
      .map((c, index) => ({ c, index, pos: startPos.get(Number(c.id ?? 0)) }))
      .sort((a, b) => {
        if (a.pos != null && b.pos != null) return a.pos - b.pos;
        if (a.pos != null) return -1;
        if (b.pos != null) return 1;
        return a.index - b.index;
      });
    for (const { c } of ordered) makeCard(c);
    const roots: ThreadCard[] = [];
    for (const { c } of ordered) {
      const card = cardOf.get(Number(c.id ?? 0));
      if (!card) continue;
      // Walk the paraIdParent chain to the thread root (bounded — a malformed
      // cyclic chain falls back to a top-level card instead of spinning).
      const paraId = paraIdOf(c);
      let parentParaId = paraId ? exByParaId.get(paraId)?.paraIdParent : undefined;
      let root: CommentEntry | undefined;
      let guard = 0;
      while (parentParaId && guard++ < 32) {
        const parent = comments.find((k) => paraIdOf(k) === parentParaId);
        if (!parent) break;
        root = parent;
        const parentEx = exByParaId.get(parentParaId);
        parentParaId = parentEx?.paraIdParent;
      }
      const rootCard = root ? cardOf.get(Number(root.id)) : undefined;
      if (rootCard && rootCard.id !== card.id) rootCard.replies.push(card);
      else roots.push(card);
    }
    pane.comments = JSON.stringify(roots);
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
