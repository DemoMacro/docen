import type { Editor } from "@docen/docx/core";
import { TextSelection } from "@tiptap/pm/state";

import { collectRevisions } from "../extensions/track-changes";

/** The revisions commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface RevisionsHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — card clicks scroll the revision into view. */
  bridge(): { scrollIntoView(pos: number): void } | undefined;
  /** The host element — the shadow-DOM root for the revisions pane. */
  element(): HTMLElement;
}

/**
 * The revision-tracking domain's pane interactions, split out of the host
 * element: syncing the reviewing pane's card list after every render (the
 * pane is a pure view of the PM revision marks), and the select/accept/reject
 * card actions (accept/reject go through the editor commands' by-id form so
 * each stays one transaction — one undo step).
 */
export class RevisionsCommands {
  constructor(private readonly host: RevisionsHost) {}

  #paneEl(): (HTMLElement & { revisions?: string; activeIndex?: string }) | null {
    return this.host.element().shadowRoot?.querySelector("docen-revisions-pane") ?? null;
  }

  /** Sync the reviewing pane's card list from the doc's revision marks after
   *  every render run. */
  syncRevisionsPane(): void {
    const pane = this.#paneEl();
    if (!pane) return;
    const editor = this.host.editor();
    if (!editor) return;
    const revisions = collectRevisions(editor.state.doc);
    pane.revisions = JSON.stringify(
      revisions.map((r, index) => ({
        index,
        type: r.type,
        author: r.author,
        date: r.date,
        text: r.text,
      })),
    );
    this.#syncActiveIndex();
  }

  /** Highlight the card whose range covers the selection (mirrors the
   *  comments pane's active card). Arrow property — it rides the editor's
   *  selectionUpdate event, so it must not lose `this`. */
  readonly syncActiveRevision = (): void => {
    this.#syncActiveIndex();
  };

  #syncActiveIndex(): void {
    const pane = this.#paneEl();
    const editor = this.host.editor();
    if (!pane || !editor) return;
    const { from, to } = editor.state.selection;
    // Same edge-inclusive overlap the accept/reject commands pick with.
    const hit = collectRevisions(editor.state.doc).findIndex((r) => r.from <= to && r.to >= from);
    pane.setAttribute("active-index", hit >= 0 ? String(hit) : "");
  }

  /** revision:select → select the revision's range and scroll it into view
   *  (Word's reviewing pane scrolls the tracked text into view on click). */
  readonly onRevisionSelect = (event: CustomEvent<{ index?: number }>): void => {
    const editor = this.host.editor();
    const index = event.detail?.index;
    if (!editor || index == null) return;
    const revision = collectRevisions(editor.state.doc)[index];
    if (!revision) return;
    editor.view.dispatch(
      editor.state.tr.setSelection(
        new TextSelection(
          editor.state.doc.resolve(revision.from),
          editor.state.doc.resolve(revision.to),
        ),
      ),
    );
    this.host.bridge()?.scrollIntoView(revision.from);
  };

  /** revision:accept / revision:reject → the by-id command form; the command
   *  owns the one-transaction accept/reject and the move-to-next walk. */
  readonly onRevisionAccept = (event: CustomEvent<{ index?: number }>): void => {
    this.#apply(event.detail?.index, "accept-change");
  };

  readonly onRevisionReject = (event: CustomEvent<{ index?: number }>): void => {
    this.#apply(event.detail?.index, "reject-change");
  };

  #apply(index: number | undefined, command: "accept-change" | "reject-change"): void {
    const editor = this.host.editor();
    if (!editor || index == null) return;
    const revision = collectRevisions(editor.state.doc)[index];
    if (!revision) return;
    (editor.commands as unknown as Record<string, (id?: string) => unknown>)[command](
      String(revision.id),
    );
  }
}
