import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { ReplaceStep } from "@tiptap/pm/transform";

/**
 * Footnote/endnote orphan cleanup — Word semantics for deleting a reference:
 * removing the last `footnoteReference`/`endnoteReference` atom of a note
 * deletes the note (word/footnotes.xml entries live in doc attrs
 * `documentExtras.footnotes`/`endnotes`, outside the PM doc, so the reference
 * and its content must be kept in sync by hand). An appendTransaction plugin
 * watches deleting transactions and prunes note bodies no longer referenced;
 * the pruned attrs ride the same dispatch, so one undo restores both the
 * reference and the note.
 */

interface DocumentExtras {
  footnotes?: Array<{ id?: number; [key: string]: unknown }>;
  endnotes?: Array<{ id?: number; [key: string]: unknown }>;
  [key: string]: unknown;
}

/** The note ids a bare `footnoteReference`/`endnoteReference` branch carries —
 *  the flat number form (`{ footnoteReference: 1 }`) or the option object
 *  form (`{ footnoteReference: { id } }`). */
export function noteRefId(
  branch: Record<string, unknown>,
): { kind: "footnoteReference" | "endnoteReference"; id?: number } | null {
  for (const key of ["footnoteReference", "endnoteReference"] as const) {
    const value = branch[key];
    if (typeof value === "number") return { kind: key, id: value };
    if (value && typeof value === "object" && typeof (value as { id?: unknown }).id === "number")
      return { kind: key, id: (value as { id: number }).id };
  }
  return null;
}

/** Every note id still referenced in the doc — one walk over the passthrough
 *  atoms (the reference carrier), keyed by the office-open branch name. */
export function liveNoteIds(doc: PMNode): {
  footnoteReference: Set<number>;
  endnoteReference: Set<number>;
} {
  const live = { footnoteReference: new Set<number>(), endnoteReference: new Set<number>() };
  doc.descendants((node) => {
    if (node.type.name !== "inlinePassthrough") return true;
    const data = node.attrs.data;
    if (typeof data !== "string") return true;
    try {
      const branch = JSON.parse(data) as Record<string, unknown>;
      const ref = noteRefId(branch);
      if (ref?.id != null) live[ref.kind as "footnoteReference" | "endnoteReference"].add(ref.id);
    } catch {
      /* malformed atom JSON — not a note reference */
    }
    return true;
  });
  return live;
}

/** The channels whose array entries lost their last reference. Entries without
 *  a numeric id are kept (nothing identifies them); every other key of
 *  documentExtras rides through untouched. */
export function prunedExtras(
  extras: DocumentExtras,
  live: { footnoteReference: Set<number>; endnoteReference: Set<number> },
): DocumentExtras | null {
  let changed: DocumentExtras | null = null;
  for (const [channel, kind] of [
    ["footnotes", "footnoteReference"],
    ["endnotes", "endnoteReference"],
  ] as const) {
    const notes = extras[channel];
    if (!Array.isArray(notes)) continue;
    const kept = notes.filter((note) => typeof note?.id !== "number" || live[kind].has(note.id));
    if (kept.length === notes.length) continue;
    changed = { ...(changed ?? extras), [channel]: kept };
  }
  return changed;
}

const notesCleanupKey = new PluginKey("notesCleanup");

export const NotesCleanup = Extension.create({
  name: "notesCleanup",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: notesCleanupKey,
        appendTransaction(transactions, _oldState, newState) {
          if (!transactions.some((tr) => tr.docChanged)) return null;
          // Only deleting transactions can orphan a note (a moved reference
          // still exists in the new doc). PM re-feeds our own trs — they carry
          // no deletion and exit here.
          const deleted = transactions.some((tr) =>
            tr.steps.some((step) => step instanceof ReplaceStep && step.from !== step.to),
          );
          if (!deleted) return null;
          const extras = (newState.doc.attrs.documentExtras ?? {}) as DocumentExtras;
          const pruned = prunedExtras(extras, liveNoteIds(newState.doc));
          if (!pruned) return null;
          return newState.tr.setDocAttribute("documentExtras", pruned);
        },
      }),
    ];
  },
});
