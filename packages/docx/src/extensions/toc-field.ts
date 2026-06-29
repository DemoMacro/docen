import { Node } from "../core";
import { attrNative } from "./utils";

/**
 * TOC field (`tocField`) — a block container representing a DOCX table of
 * contents, with the rendered entries as editable `content` and the TOC field
 * switches on `attrs.options`.
 *
 * Named `tocField` (not `tableOfContents`) to avoid colliding with the official
 * `@tiptap/extension-table-of-contents`, which is an `Extension` (a live outline
 * generator injected into the editor) that already owns the `tableOfContents`
 * name. Same-name extensions dedupe in Tiptap, and since the official one is not
 * a node it would erase this node type from the schema. The two coexist for
 * different purposes: this node persists a DOCX's rendered TOC; the official
 * extension drives the heading-outline pane.
 *
 * Structuring the TOC as a node (instead of opaque passthrough) is what fixes
 * the export crash. Each entry paragraph's `w:hyperlink` wraps a HYPERLINK field
 * whose content-less runs (fldChar begin/separate/end) office-open parses as
 * `null`. As opaque passthrough those nulls survived verbatim to
 * `generateDocument`, where office-open's `stringifyRunInline(null).break`
 * crashed. Resolving the entries through `resolveParagraphChildren` drops the
 * nulls (the existing `child !== null` guard), so compile rebuilds clean entries
 * and the generate path never sees a null — no office-open change required.
 *
 * DOCX serialization is inlined in DocxManager (resolve/compile read/write
 * `attrs.options` + the entry content directly), so no renderDocx/parseDocx is
 * needed here — the same pattern as the details extension.
 */
export const TocField = Node.create({
  name: "tocField",
  group: "block",
  content: "block+",

  addAttributes() {
    return {
      // TOC field switches (hyperlink, headingStyleRange, …); carried verbatim,
      // never rendered to HTML.
      options: attrNative(),
    };
  },

  parseHTML() {
    return [{ tag: "div.docx-toc" }];
  },

  renderHTML() {
    return ["div", { class: "docx-toc" }, 0];
  },
});
