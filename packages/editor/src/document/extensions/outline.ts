import { Extension } from "@docen/docx/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";

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

/**
 * Read-only navigation-outline generator (replaces @tiptap/extension-table-of-
 * contents).
 *
 * Walks the doc for headings on every doc change and emits the anchor list via
 * `onUpdate`. Crucially it does NO `setNodeMarkup`, so no content re-validation:
 * the official extension's appendTransaction calls setNodeMarkup to inject
 * id/data-toc-id on each heading, and on large / list-rich docs that
 * re-validation throws "Invalid content for node listItem" — aborting the
 * paginator's reflow transaction so the whole document piles on page 1. Outline
 * clicks jump by `pos`, so the injected ids aren't needed; a read-only walk
 * suffices. The `id` is a heading-order index, stable across re-flows (which
 * only repaginate, never reorder headings).
 */
export const Outline = Extension.create<{ onUpdate: (anchors: readonly OutlineAnchor[]) => void }>({
  name: "docenOutline",
  addProseMirrorPlugins() {
    const { onUpdate } = this.options;
    return [
      new Plugin({
        key: new PluginKey("docenOutline"),
        view(view) {
          const emit = (): void => {
            const anchors: OutlineAnchor[] = [];
            let idx = 0;
            view.state.doc.descendants((node, pos) => {
              if (node.type.name === "heading" && node.textContent.length > 0) {
                anchors.push({
                  id: "h" + idx,
                  pos,
                  textContent: node.textContent,
                  originalLevel: (node.attrs.level as number) ?? 1,
                });
                idx++;
                return false;
              }
              return true;
            });
            onUpdate(anchors);
          };
          emit();
          return {
            update(v, prevState) {
              if (v.state.doc !== prevState.doc) emit();
            },
          };
        },
      }),
    ];
  },
});
