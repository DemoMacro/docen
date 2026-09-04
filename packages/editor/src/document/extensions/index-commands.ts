import { type StylesOptions } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Fragment, Slice } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import type { PageOf } from "./toc";

/**
 * Index commands — the ribbon References tab's Insert Index and Update Index
 * buttons (the Mark Entry interaction itself lives in the host: it prompts).
 *
 * Marking an entry seeds an `XE "…"` field (a cached-less fldSimple, which the
 * projection renders as nothing — Word's invisible index marker). These
 * commands collect the XE fields and emit Word's index shape: one Index1
 * paragraph per main entry, Index2 for the `主:子` sub-entries, a right
 * dotted-leader tab, and the marked paragraph's live page number from the
 * canvas caret map (the same `pageOf` bridge the TOC uses).
 *
 * Entries ride plain Index-styled paragraphs rather than an INDEX-field
 * container: the field result would have to span paragraphs, which the inline
 * field model can't carry; Update Index provides the rebuild instead.
 */

/** Word's built-in index styles indent 220 twips per level — the same ladder
 *  the TOC styles use. */
const INDEX_INDENT_TW = 220;

/** One collected index entry: the entry text (main[:sub]) and the pages its
 *  XE fields sit on. */
interface IndexEntry {
  text: string;
  pages: number[];
}

/** Collect the document's XE fields, grouped by entry text. Pages dedupe and
 *  stay in document order. */
function collectEntries(doc: PMNode, pageOf?: PageOf): Map<string, number[]> {
  const entries = new Map<string, number[]>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const page = pageOf?.(pos + 1);
    node.forEach((child) => {
      if (child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        const instr = (data.simpleField?.instruction ?? "").trim();
        const m = /^XE\s+"([^"]*)"/i.exec(instr);
        if (!m?.[1]) return;
        const pages = entries.get(m[1]) ?? [];
        if (typeof page === "number" && !pages.includes(page)) pages.push(page);
        entries.set(m[1], pages);
      } catch {
        // opaque verbatim blobs without field data — skip
      }
    });
    return true;
  });
  return entries;
}

/** Index entry paragraphs from the collected XE fields: mains sorted with the
 *  document language's collation, each `主:子` sub-entry nested under its
 *  main. */
function buildIndexParagraphs(
  entries: Map<string, number[]>,
  tabPositionTw = 9350,
): { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] {
  const groups = new Map<string, Map<string, number[]>>();
  for (const [raw, pages] of entries) {
    const sep = raw.indexOf(":");
    const main = sep < 0 ? raw : raw.slice(0, sep);
    const sub = sep < 0 ? "" : raw.slice(sep + 1);
    const subs = groups.get(main) ?? new Map<string, number[]>();
    if (!groups.has(main)) groups.set(main, subs);
    const merged = subs.get(sub) ?? [];
    for (const page of pages) if (!merged.includes(page)) merged.push(page);
    subs.set(sub, merged);
  }
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];
  const entry = (text: string, level: number, pages: number[]) => ({
    type: "paragraph",
    attrs: {
      style: `Index${level + 1}`,
      tabStops: [{ type: "right", position: tabPositionTw, leader: "dot" }],
      ...(level > 0 ? { indent: { left: level * INDEX_INDENT_TW } } : {}),
    },
    content: [
      { type: "text", text },
      { type: "tab" },
      ...(pages.length > 0 ? [{ type: "text", text: pages.join(", ") }] : []),
    ],
  });
  for (const main of [...groups.keys()].sort((a, b) => a.localeCompare(b, "zh"))) {
    const subs = groups.get(main)!;
    const own = subs.get("") ?? [];
    out.push(entry(main, 0, own));
    for (const sub of [...subs.keys()]
      .filter((s) => s !== "")
      .sort((a, b) => a.localeCompare(b, "zh"))) {
      out.push(entry(sub, 1, subs.get(sub)!));
    }
  }
  return out;
}

/** The first Index-styled paragraph — an existing index's head, the spot the
 *  rebuild replaces from. */
function findIndexHead(doc: PMNode): number | null {
  let pos: number | null = null;
  doc.descendants((node, at) => {
    if (pos != null) return false;
    if (node.type.name === "paragraph" && /^Index\d$/.test(String(node.attrs.style ?? ""))) {
      pos = at;
      return false;
    }
    return true;
  });
  return pos;
}

/** Stamp the Index1/Index2 style definitions when the document carries none
 *  (compile passes doc.attrs.styles straight through). */
function stampIndexStyles(tr: Transaction): void {
  const styles = { ...((tr.doc.attrs.styles ?? {}) as Record<string, unknown>) };
  const paragraphStyles = (styles.paragraphStyles ?? []) as {
    id?: string;
    name?: string;
    basedOn?: string;
    next?: string;
    indent?: { left?: number };
  }[];
  const missing = (["Index1", "Index2"] as const).filter(
    (id) => !paragraphStyles.some((style) => style.id === id),
  );
  if (missing.length === 0) return;
  for (const id of missing) {
    const level = Number(id.slice(-1)) - 1;
    paragraphStyles.push({
      id,
      name: `index ${id.slice(-1)}`,
      basedOn: "Normal",
      next: id,
      ...(level > 0 ? { indent: { left: level * INDEX_INDENT_TW } } : {}),
    });
  }
  tr.step(new DocAttrStep("styles", { ...styles, paragraphStyles }));
}

export const IndexCommands = Extension.create({
  name: "docenIndexCommands",

  addCommands() {
    return {
      // Insert a fresh index at the caret from the document's XE fields.
      // PM's nodeFromJSON keeps the command DOM-free (the viewless editor and
      // the headless tests have no window).
      "insert-index":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const paragraphs = buildIndexParagraphs(collectEntries(state.doc, pageOf), tabPositionTw);
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampIndexStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            tr.replaceSelection(new Slice(Fragment.fromArray(nodes), 0, 0)).scrollIntoView();
          }
          return true;
        },
      // Rebuild the index in place — Word's F9. The Index-styled paragraphs
      // are replaced from the first one's position; XE fields elsewhere are
      // untouched.
      "update-index":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const head = findIndexHead(state.doc);
          if (head == null) return false;
          const paragraphs = buildIndexParagraphs(collectEntries(state.doc, pageOf), tabPositionTw);
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampIndexStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            // Drop every Index-styled paragraph, then splice the fresh block
            // in at the head's (unchanged — earlier) position.
            const doomed: { from: number; to: number }[] = [];
            state.doc.descendants((node, at) => {
              if (
                node.type.name === "paragraph" &&
                /^Index\d$/.test(String(node.attrs.style ?? ""))
              )
                doomed.push({ from: at, to: at + node.nodeSize });
              return true;
            });
            for (const range of doomed.reverse()) tr.delete(range.from, range.to);
            tr.insert(head, nodes);
          }
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenIndexCommands: {
      "insert-index": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
      "update-index": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
    };
  }
}
