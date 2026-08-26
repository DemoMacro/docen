import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import { Extension, type Editor } from "@docen/docx/core";

/** A heading anchor for the navigation outline. `id` is a stable heading-order
 *  index (so the outline's signature dedups without rebuilding the tree on
 *  every re-flow); `pos` is live (outline clicks jump to it — see
 *  DocenDocument.#onOutlineSelect); `textContent` + `originalLevel` drive the
 *  nested tree. */
export interface OutlineAnchor {
  id: string;
  pos: number;
  textContent: string;
  originalLevel: number;
}

/** Collect heading anchors: a heading IS a paragraph, so the walk resolves each
 *  paragraph's level via detectHeadingLevel — the lifted HeadingLevel attr, an
 *  explicit outlineLevel, or a pStyle naming a heading style (directly /
 *  localized name / basedOn chain). Read-only: no setNodeMarkup means no
 *  content re-validation (the official TOC extension's injected ids
 *  re-validate and can abort the reflow transaction on list-rich docs);
 *  outline clicks jump by `pos`, so ids aren't needed. `id` is a heading-order
 *  index, stable across re-flows (which only repaginate, never reorder). */
function collectAnchors(editor: Editor): OutlineAnchor[] {
  const anchors: OutlineAnchor[] = [];
  let idx = 0;
  // Styles snapshot once per walk: detectHeadingLevel indexes it for
  // style-based detection, so re-indexing per paragraph is O(n·m) otherwise.
  const styles = (editor.state.doc.attrs as { styles?: StylesOptions }).styles;
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const level = detectHeadingLevel(
      {
        heading: (node.attrs.heading as string) || undefined,
        style: (node.attrs.style as string) || undefined,
        outlineLevel: node.attrs.outlineLevel as number | undefined,
      },
      styles,
    );
    if (level != null && node.textContent.length > 0) {
      anchors.push({
        id: "h" + idx,
        // Inside the paragraph — descendants' pos is the node's start, one
        // shy of the content; caretRect rejects block-boundary positions.
        pos: pos + 1,
        textContent: node.textContent,
        originalLevel: level,
      });
      idx++;
      return false;
    }
    return true;
  });
  return anchors;
}

/**
 * Read-only navigation-outline generator (replaces @tiptap/extension-table-of-
 * contents). Emits the anchor list via `onUpdate` on create and on every doc
 * change. Lifecycle-driven, NOT a plugin view: the canvas route runs a
 * viewless editor (element: null) whose plugins are registered by hand —
 * plugin views never install, so a Plugin({ view }) factory would never run.
 */
export const Outline = Extension.create<{ onUpdate: (anchors: readonly OutlineAnchor[]) => void }>({
  name: "docxOutline",
  onCreate() {
    this.options.onUpdate(collectAnchors(this.editor));
  },
  onTransaction({ transaction }) {
    if (transaction.docChanged) {
      this.options.onUpdate(collectAnchors(this.editor));
    }
  },
});
