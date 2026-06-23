import { Extension } from "@docen/docx/core";

/**
 * Paragraph/heading split support for C-route pagination — mirrors split-table.
 *
 * A paragraph taller than the remaining page space splits across pages: the
 * paginator cuts it at a line boundary into a head (first N lines, current
 * page) and a tail (rest, next page). Both halves share a `splitGroup` id;
 * `unwrapPages` (in utils/merge.ts) merges them back into one paragraph on
 * export, so the split is editor-only and round-trip-transparent. `splitPart`
 * marks head vs tail.
 *
 * Headings split the same way (they share ParagraphPropertiesOptionsBase with
 * paragraphs, so the same attrs apply).
 *
 * `splitGroup`/`splitPart` are injected as GLOBAL attributes on paragraph +
 * heading via `addGlobalAttributes` (the SplitMarks Extension below), NOT by
 * same-name `.extend({ addAttributes })` overrides. The previous override
 * approach broke Tiptap's parent chain — `this.parent?.()` resolved to the base
 * Tiptap node instead of @docen/docx's Heading/Paragraph, so the `styleId`
 * attribute lost its `parseHTML` in the editor schema (only `{ default }`
 * survived), which broke HTML-paste styleId parsing. Global attributes APPEND to
 * whatever the node's own `addAttributes` already declares, so the docx nodes'
 * `styleId`/`parseHTML` stays intact.
 *
 * See CLAUDE.md → Pagination Architecture (C-route) and CONTRIBUTING.md.
 */

/** Injects the editor-only `splitGroup`/`splitPart` attrs onto paragraph +
 *  heading for the paginator. Registered as an Extension (global attrs), not as
 *  node overrides — see the module doc above for why. */
export const SplitMarks = Extension.create({
  name: "splitMarks",
  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          // Split id shared by the head + tail of one original paragraph. null
          // on an un-split paragraph. Editor-only — cleared on export. Mirrors
          // table `splitGroup`.
          splitGroup: { default: null, parseHTML: () => null, rendered: false },
          splitPart: { default: null, parseHTML: () => null, rendered: false },
        },
      },
    ];
  },
});
