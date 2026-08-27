import { Extension, type AnyExtension, type Editor } from "@docen/docx/core";
import { CharacterCount, TrailingNode } from "@tiptap/extensions";
import { search } from "prosemirror-search";

import { type DocenAddin, type DocenHost } from "../ui";
import { DocumentCommands } from "./extensions/commands";
import { Outline, type OutlineAnchor } from "./extensions/outline";
import { TrackChanges } from "./extensions/track-changes";

/** A `<docen-document>` host: a {@link DocenHost} carrying a Tiptap `Editor`. */
export type DocumentHost = DocenHost<Editor>;

/** A docen-document add-in. Extends the editor-agnostic {@link DocenAddin} with a
 *  Tiptap `extensions` contribution — document-specific, since presentation
 *  (LeaferJS) and workbook (RevoGrid) will use different engines, so engine
 *  extensions stay out of the editor-agnostic base contract. */
export interface DocumentAddin extends DocenAddin<DocumentHost> {
  readonly extensions?: readonly AnyExtension[];
}

/** Search plugin wrapper (prosemirror-search). The host listens for
 *  `navigation:search` and dispatches `setSearchState`; this extension only
 *  registers the plugin, so it has no host coupling. */
const Search = Extension.create({
  name: "docenSearch",
  addProseMirrorPlugins() {
    return [search()];
  },
});

/** Office-style word count: each CJK character counts as one, non-CJK runs split
 *  on whitespace — matches Word for mixed CJK/Latin (the default split(' ').length
 *  counts a whole CJK paragraph as a single word). */
const wordCounter = (text: string): number => {
  const cjkRe = /[一-鿿぀-ヿ가-힯]/g;
  const cjk = (text.match(cjkRe) ?? []).length;
  const western = text.replace(cjkRe, " ").split(/\s+/).filter(Boolean).length;
  return cjk + western;
};

/** Count characters by grapheme cluster so emoji / combining marks / surrogate
 *  pairs count as one (default text.length undercounts them). */
const textCounter = (text: string): number => {
  const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
  let n = 0;
  for (const _ of seg.segment(text)) n++;
  return n;
};

/** Build the default document engine extensions for the canvas route —
 *  state-based only (no view plugins: the canvas paints, the viewless editor
 *  edits). Outline reports the heading anchor list to `<docen-outline>`, so
 *  the factory takes that callback rather than capturing `this`. UndoRedo is
 *  layered by the edit bridge itself (it also registers the plugins a
 *  viewless editor skips). */
export function createDocumentExtensions(opts: {
  onOutlineUpdate: (anchors: readonly OutlineAnchor[]) => void;
}): readonly AnyExtension[] {
  return [
    // Outline: a read-only heading walk that reports the anchor list to
    // <docen-outline>.
    Outline.configure({ onUpdate: opts.onOutlineUpdate }),
    Search,
    TrailingNode,
    CharacterCount.configure({ wordCounter, textCounter }),
    // Review tab revision tracking: the toggle + live edit marking, and the
    // accept/reject/navigation commands the ribbon buttons dispatch.
    TrackChanges,
    // Ribbon commands as native Tiptap commands (editor.commands.<event>), so
    // #onCommand routes event → editor.commands[event](value) with no mapping
    // layer. Includes editor.can() for precise ribbon greying.
    DocumentCommands,
  ];
}

/** The default document add-in: engine essentials (extensions) — including
 *  DocumentCommands, which exposes every ribbon command as a native
 *  `editor.commands[name]`. Ribbon and task-pane contributions are layered on
 *  once the host renders from the merged schema. */
export function createDefaultAddin(opts: {
  onOutlineUpdate: (anchors: readonly OutlineAnchor[]) => void;
}): DocumentAddin {
  return {
    id: "docen-document",
    name: "Document",
    extensions: createDocumentExtensions(opts),
  };
}
