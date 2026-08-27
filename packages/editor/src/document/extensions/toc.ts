import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * Table of contents commands — the ribbon References tab's TOC and Update
 * Table buttons.
 *
 * The `tocField` node (engine) already round-trips a DOCX TOC's rendered
 * entries; these commands close the loop on the authoring side: scan the doc's
 * headings (the outline walk's rule), emit one entry paragraph per heading —
 * TOC1-3 style, the Word built-in per-level indent, a right dotted-leader tab,
 * and the heading's live page number from the canvas caret map — and insert or
 * refresh the `tocField` at the caret.
 *
 * Page numbers come from the host's bridge (`pageOf`): pagination lives in the
 * canvas layout, which the headless command chain can't see. The callback is a
 * command argument for the same reason — the extension is created before the
 * bridge exists, so the host wires it per dispatch (see #onCommand's local
 * branch). Inserting a TOC repaginates the doc, so the host re-runs the update
 * once the fresh layout lands; Word behaves the same (insert shows stale
 * numbers until the field updates).
 */

/** A heading's page lookup, wired by the host from the canvas caret map. */
export type PageOf = (pos: number) => number | null | undefined;

/** Word's built-in TOC styles indent 220 twips per level (TOC2 at 220, TOC3
 *  at 440). Stamped directly on the entry so the hierarchy reads even when the
 *  doc carries no TOC1-3 style definitions. */
const TOC_INDENT_TW = 220;

/** Entry paragraphs for the headings the default TOC covers (levels 1-3). */
function buildTocEntries(
  doc: PMNode,
  pageOf?: PageOf,
  tabPositionTw = 9350,
): { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] {
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const level = detectHeadingLevel(
      {
        heading: (node.attrs.heading as string) || undefined,
        style: (node.attrs.style as string) || undefined,
        outlineLevel: node.attrs.outlineLevel as number | undefined,
      },
      styles,
    );
    if (level == null || level > 3 || node.textContent.length === 0) return true;
    const page = pageOf?.(pos + 1);
    out.push({
      type: "paragraph",
      attrs: {
        style: `TOC${level}`,
        tabStops: [{ type: "right", position: tabPositionTw, leader: "dot" }],
        ...(level > 1 ? { indent: { left: (level - 1) * TOC_INDENT_TW } } : {}),
      },
      content: [
        {
          type: "text",
          text: node.textContent,
          marks: [{ type: "link", attrs: { href: `#_Toc${out.length + 1}` } }],
        },
        { type: "tab" },
        // A blank page (unmapped heading) omits the number run — an empty
        // text node is illegal in PM.
        ...(typeof page === "number" ? [{ type: "text", text: String(page) }] : []),
      ],
    });
    return true;
  });
  return out;
}

/** The first tocField in the doc, with its position (null when none). */
function findTocField(doc: PMNode): { node: PMNode; pos: number } | null {
  let found: { node: PMNode; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name === "tocField") {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

export const TocCommands = Extension.create({
  name: "docenTocCommands",

  addCommands() {
    return {
      // Insert a fresh TOC (Word's default: heading levels 1-3, hyperlinked)
      // at the caret. Entries build from the live headings; page numbers ride
      // the host's pageOf (null → blank until the post-insert update lands).
      // PM's nodeFromJSON (not Tiptap's insertContentAt) keeps the command
      // DOM-free — the viewless editor and the headless tests have no window.
      toc:
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, dispatch }) => {
          const entries = buildTocEntries(state.doc, pageOf, tabPositionTw);
          if (entries.length === 0) return false;
          const node = state.schema.nodeFromJSON({
            type: "tocField",
            attrs: { options: { headingStyleRange: "1-3", hyperlink: true } },
            content: entries,
          });
          if (!node) return false;
          if (dispatch) dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
          return true;
        },
      // Rebuild the first TOC's entries from the current headings — Word's F9.
      // The field switches (attrs.options) are preserved verbatim. Entries
      // materialize through nodeFromJSON — replaceWith takes PM nodes, not
      // JSON (JSON would silently land outside the field).
      "update-toc":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const found = findTocField(state.doc);
          if (!found) return false;
          const entries = buildTocEntries(state.doc, pageOf, tabPositionTw);
          if (entries.length === 0) return false;
          if (dispatch) {
            const nodes = entries.map((entry) => state.schema.nodeFromJSON(entry));
            tr.replaceWith(found.pos + 1, found.pos + found.node.nodeSize - 1, nodes);
          }
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenTocCommands: {
      toc: (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
      "update-toc": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
    };
  }
}
