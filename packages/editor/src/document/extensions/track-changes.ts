import { Extension } from "@docen/docx/core";
import type { Node as PMNode, MarkType } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import { ReplaceStep } from "@tiptap/pm/transform";

/**
 * Track Changes — Word revision tracking for the viewless canvas route.
 *
 * The `insertion`/`deletion` marks (engine extensions) already round-trip
 * w:ins/w:del and render struck/underlined; this extension adds the missing
 * workflow: a tracking toggle, marking of live edits, and accept/reject +
 * navigation commands. All five names are the ribbon Review tab's event
 * attributes (`editor.commands.<event>`), greyed until registered in
 * WIRED_DISPATCH.
 *
 * Marking scope is TEXT EDITS INSIDE ONE PARAGRAPH — the same inline scope the
 * OOXML round-trip covers (office-open parses inline w:ins/w:del only).
 * Structural edits (Enter, paste with block content, node deletion) apply
 * untracked, and undo/redo replays as-is — tracking an undo would re-mark the
 * reverted edit and corrupt the revision list.
 *
 * Deletions keep the text (Word: struck, removed on accept): an appendTransaction
 * plugin re-inserts the removed runs from the pre-edit doc — preserving their
 * own rPr marks — under a `deletion` mark placed AFTER the inserted text
 * (LibreOffice's replacement order). Deleting already-struck text is a no-op
 * (Word refuses to delete a deletion); the restore keeps it struck.
 */

/** Revision identity stamped on tracked edits (w:ins/@w:author shape). One
 *  author until a user-name setting exists; Word's "By author" palette starts
 *  at red, which the canvas projection stamps for display. */
const REVISION_AUTHOR = "docen";

const revisionDate = (): string =>
  // Word writes second precision ("2026-08-28T09:30:00Z").
  new Date().toISOString().replace(/\.\d+Z$/, "Z");

interface RevisionAttrs {
  id: number;
  author: string;
  date: string;
}

const isRevisionMark = (name: string): boolean => name === "insertion" || name === "deletion";

/** Highest existing revision id + 1 — w:id is a document-unique integer. */
function nextRevisionId(doc: PMNode): number {
  let max = 0;
  doc.descendants((node) => {
    if (!node.isText) return true;
    for (const mark of node.marks) {
      if (!isRevisionMark(mark.type.name)) continue;
      const id = (mark.attrs as { id?: unknown }).id;
      if (typeof id === "number" && id > max) max = id;
    }
    return true;
  });
  return max + 1;
}

/** A tracked edit's metadata. Consecutive same-author edits merge into one
 *  record (Word): text touching a record by this author reuses its attrs
 *  instead of allocating a fresh id — typing a sentence makes ONE w:ins. */
function revisionAttrs(doc: PMNode, from: number, types: MarkType[]): RevisionAttrs {
  const touch = (node: PMNode | null): RevisionAttrs | null => {
    if (!node?.isText) return null;
    for (const mark of node.marks) {
      if (!types.includes(mark.type)) continue;
      if ((mark.attrs as { author?: string }).author !== REVISION_AUTHOR) continue;
      return mark.attrs as unknown as RevisionAttrs;
    }
    return null;
  };
  const $pos = doc.resolve(from);
  return (
    touch($pos.nodeBefore) ??
    touch($pos.nodeAfter) ?? {
      id: nextRevisionId(doc),
      author: REVISION_AUTHOR,
      date: revisionDate(),
    }
  );
}

/** Plugin state = tracking on/off; the toggle command rides a meta so the
 *  flag changes with the transaction that turned it on. */
const trackChangesKey = new PluginKey<boolean>("docenTrackChanges");

/** Review commands (accept/reject) deliberately delete revision text — their
 *  transactions carry this meta so the marking plugin never "restores" what
 *  the review just removed. */
const skipTrackingKey = new PluginKey<boolean>("docenTrackChangesSkip");

const trackChangesPlugin = new Plugin<boolean>({
  key: trackChangesKey,
  state: {
    init: () => false,
    apply: (tr, value) => tr.getMeta(trackChangesKey) ?? value,
  },
  appendTransaction(transactions, oldState, newState) {
    if (!trackChangesKey.getState(newState)) return null;
    // Rounds re-fed by PM (its "appendedTransaction" tag marks our own trs),
    // the toggle itself, review commands, and undo/redo replay untouched.
    if (
      transactions.some(
        (tr) =>
          tr.getMeta("appendedTransaction") !== undefined ||
          tr.getMeta(trackChangesKey) !== undefined ||
          tr.getMeta(skipTrackingKey) !== undefined ||
          tr.getMeta("history$") !== undefined,
      )
    ) {
      return null;
    }
    const insertionType = newState.schema.marks.insertion;
    const deletionType = newState.schema.marks.deletion;
    if (!insertionType || !deletionType) return null;

    // One text ReplaceStep per transaction is the tracked shape; anything
    // else (multi-step commands, block edits) applies untracked.
    const steps = transactions.flatMap((tr) => tr.steps);
    if (steps.length !== 1 || !(steps[0] instanceof ReplaceStep)) return null;
    const step = steps[0];
    const { from, to, slice } = step;

    // The inserted side must be plain text: a single text node, no open
    // depths (structure would need block-level revisions to track).
    const inserted = slice.content;
    let insertedText = "";
    if (slice.openStart !== 0 || slice.openEnd !== 0) return null;
    if (inserted.childCount === 1 && inserted.firstChild!.isText) {
      insertedText = inserted.firstChild!.text ?? "";
    } else if (inserted.childCount !== 0) {
      return null;
    }

    // The removed side must be plain text inside one paragraph of the OLD doc.
    const $from = oldState.doc.resolve(from);
    const $to = oldState.doc.resolve(to);
    if (!$from.sameParent($to) || !$from.parent.isTextblock) return null;
    const removedNodes: PMNode[] = [];
    if (to > from) oldState.doc.slice(from, to).content.forEach((node) => removedNodes.push(node));
    for (const node of removedNodes) {
      if (!node.isText || node.text?.includes("\n")) return null;
    }
    if (!insertedText && removedNodes.length === 0) return null;

    const tr = newState.tr;
    const removedAlreadyStruck = removedNodes.every((node) =>
      node.marks.some((m) => m.type === deletionType),
    );

    // Inserted text → insertion mark, reusing a touching record.
    if (insertedText) {
      tr.addMark(
        from,
        from + insertedText.length,
        insertionType.create(revisionAttrs(newState.doc, from, [insertionType])),
      );
    }
    // Removed text stays: re-insert the pre-edit runs (original rPr marks
    // preserved) struck, after the inserted text. Text that was ALREADY struck
    // keeps its record — the delete is refused, the restore just keeps it.
    if (removedNodes.length > 0) {
      const at = from + insertedText.length;
      const record = revisionAttrs(newState.doc, at, [deletionType]);
      const marked = removedAlreadyStruck
        ? removedNodes
        : removedNodes.map((node) =>
            node.marks.some((m) => m.type === deletionType)
              ? node
              : node.mark([...node.marks, deletionType.create(record)]),
          );
      tr.insert(at, marked);
    }
    return tr;
  },
});

/** Contiguous revision runs of one record, merged for accept/reject picking. */
interface RevisionRange {
  from: number;
  to: number;
  type: "insertion" | "deletion";
  id: unknown;
}

function revisionRanges(doc: PMNode): RevisionRange[] {
  const out: RevisionRange[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const mark = node.marks.find((m) => isRevisionMark(m.type.name));
    if (!mark) return true;
    const from = pos;
    const to = pos + node.nodeSize;
    const last = out[out.length - 1];
    const id = (mark.attrs as { id?: unknown }).id;
    if (last && last.type === mark.type.name && last.id === id && last.to === from) {
      last.to = to;
    } else {
      out.push({ from, to, type: mark.type.name as "insertion" | "deletion", id });
    }
    return true;
  });
  return out;
}

/** The revision a command acts on: the one overlapping the selection (an empty
 *  selection inside one counts), else the first after it — Word moves forward
 *  without wrapping. */
function pickRevision(ranges: RevisionRange[], from: number, to: number): RevisionRange | null {
  return (
    ranges.find((r) => r.from <= to && r.to >= from) ?? ranges.find((r) => r.from >= to) ?? null
  );
}

export const TrackChanges = Extension.create({
  name: "docenTrackChanges",

  addProseMirrorPlugins() {
    return [trackChangesPlugin];
  },

  addCommands() {
    return {
      "track-changes":
        (enabled?: boolean) =>
        ({ state, tr, dispatch }) => {
          const current = trackChangesKey.getState(state) ?? false;
          const next = typeof enabled === "boolean" ? enabled : !current;
          if (next === current) return true;
          if (dispatch) tr.setMeta(trackChangesKey, next);
          return true;
        },
      "accept-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = pickRevision(
            revisionRanges(state.doc),
            state.selection.from,
            state.selection.to,
          );
          if (!target) return false;
          if (!dispatch) return true;
          if (target.type === "insertion")
            tr.removeMark(target.from, target.to, state.schema.marks.insertion!);
          else tr.delete(target.from, target.to);
          tr.setMeta(skipTrackingKey, true);
          tr.setSelection(TextSelection.near(tr.doc.resolve(target.from)));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "reject-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = pickRevision(
            revisionRanges(state.doc),
            state.selection.from,
            state.selection.to,
          );
          if (!target) return false;
          if (!dispatch) return true;
          if (target.type === "deletion")
            tr.removeMark(target.from, target.to, state.schema.marks.deletion!);
          else tr.delete(target.from, target.to);
          tr.setMeta(skipTrackingKey, true);
          tr.setSelection(TextSelection.near(tr.doc.resolve(target.from)));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "previous-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = revisionRanges(state.doc)
            .filter((r) => r.to < state.selection.from)
            .pop();
          if (!target) return false;
          if (!dispatch) return true;
          tr.setSelection(TextSelection.create(tr.doc, target.from));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "next-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = revisionRanges(state.doc).find((r) => r.from > state.selection.to);
          if (!target) return false;
          if (!dispatch) return true;
          tr.setSelection(TextSelection.create(tr.doc, target.from));
          dispatch(tr.scrollIntoView());
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenTrackChanges: {
      "track-changes": (enabled?: boolean) => ReturnType;
      "accept-change": () => ReturnType;
      "reject-change": () => ReturnType;
      "previous-change": () => ReturnType;
      "next-change": () => ReturnType;
    };
  }
}
