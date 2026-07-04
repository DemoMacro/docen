import { Extension, type AnyExtension, type Editor } from "@docen/docx/core";
import { ListKeymap } from "@tiptap/extension-list";
import {
  CharacterCount,
  Dropcursor,
  Focus,
  Gapcursor,
  Placeholder,
  Selection,
  TrailingNode,
  UndoRedo,
} from "@tiptap/extensions";
import { search } from "prosemirror-search";

import { t, type DocenAddin, type DocenHost } from "../ui";
import { DocumentCommands } from "./extensions/commands";
import { FontMetricDecoration } from "./extensions/font-metric";
import { ImageCap } from "./extensions/image-cap";
import { DocenKeymap } from "./extensions/keymap";
import { Outline, type OutlineAnchor } from "./extensions/outline";
import { PageBreakView } from "./extensions/page-break";
import { Page, PageDocument } from "./extensions/page-node";
import { PagePlugin } from "./extensions/page-plugin";
import { SectionBreakMarks } from "./extensions/section-break";
import { SplitMarks } from "./extensions/split-paragraph";
import { SplitTable, SplitTableRow } from "./extensions/split-table";
import { WpsShapeView } from "./extensions/wps-shape-view";

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

/** Word-style word count: each CJK character counts as one, non-CJK runs split
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

/** Build the default document engine extensions. Outline reports the heading
 *  anchor list to `<docen-outline>`, so the factory takes that callback rather
 *  than capturing `this` — keeping the extensions host-agnostic and reusable. */
export function createDocumentExtensions(opts: {
  onOutlineUpdate: (anchors: readonly OutlineAnchor[]) => void;
}): readonly AnyExtension[] {
  return [
    PageDocument,
    Page,
    // ImageCap scales over-wide images to the section content width (Word
    // behavior) and runs before PagePlugin so the reflow measures the capped
    // dimensions, not the pre-cap overflow.
    ImageCap,
    PagePlugin,
    FontMetricDecoration,
    SplitTable,
    SplitTableRow,
    // Paragraph/heading split support: editor-only splitGroup/splitPart attrs
    // so the paginator can split a tall paragraph across pages at a line
    // boundary. Both halves share the splitGroup id; unwrapPages merges them.
    SplitMarks,
    // Outline: a read-only heading walk that reports the anchor list to
    // <docen-outline> (replaces @tiptap/extension-table-of-contents, whose
    // setNodeMarkup aborted the reflow on large docs).
    Outline.configure({ onUpdate: opts.onOutlineUpdate }),
    Search,
    // pageBreak NodeView — Fluent divider with a centered label while
    // show-marks is on. Schema comes from the engine's PageBreak.
    PageBreakView,
    // wpsShape NodeView — editable floating text box as two elements (outer
    // placement/rotation, inner contentDOM). Schema from the engine's WpsShape.
    WpsShapeView,
    // Centralized MS Office editing keymap (Ctrl+Enter page break, etc.) —
    // see extensions/keymap.ts. Outranks HardBreak via priority.
    DocenKeymap,
    // sectionBreak widget — a section boundary is paragraph attrs (not a
    // node), so a widget decoration paints the Fluent divider after each
    // section-carrying paragraph.
    SectionBreakMarks,
    Placeholder.configure({
      // A function (not a string) so the prompt re-reads the active locale each
      // time the decoration set is rebuilt — a locale switch refreshes the text.
      placeholder: () => t("editor.placeholder"),
      // C-route nests paragraphs inside a non-textblock page node (doc > page >
      // p). Placeholder's walk returns false at the page unless includeChildren
      // is set, so the prompt never reaches the paragraphs.
      includeChildren: true,
    }),
    // Editing-behavior set — the engine carries schema only.
    UndoRedo,
    Dropcursor,
    Gapcursor,
    TrailingNode,
    ListKeymap,
    CharacterCount.configure({ wordCounter, textCounter }),
    Focus,
    Selection,
    // Ribbon commands as native Tiptap commands (editor.commands.<event>), so
    // #onCommand routes event → editor.chain().focus()[event](value).run() with
    // no mapping layer. Includes editor.can() for precise ribbon greying.
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
