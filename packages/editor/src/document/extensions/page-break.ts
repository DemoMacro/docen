import { PageBreak } from "@docen/docx";

import { t } from "../../ui";

/**
 * Editor-only rendering of the pageBreak node as a Fluent divider with a
 * centered label (Word's "Page Break" marker), shown only while formatting
 * marks are on. The label is non-editable text.
 *
 * EXTENDS — does not re-create — the engine's PageBreak so its schema
 * (inline atom/group), parseHTML, renderHTML, and `setPageBreak` command are
 * all inherited; only the editor-time NodeView is overridden (keeps @docen/docx
 * UI-free). A same-name `Node.create` here would replace the engine node and
 * silently drop the command (`setPageBreak is not a function`).
 *
 * `ignoreMutation` tells ProseMirror to ignore DOM mutations inside the
 * NodeView (the Web Component's shadowDOM, label re-text on locale change) so
 * they don't trigger spurious view updates. The pageBreak atom carries no
 * editable content, so there is nothing for ProseMirror to track.
 */

export const PageBreakView = PageBreak.extend({
  addNodeView() {
    return () => {
      const dom = document.createElement("span");
      dom.setAttribute("data-type", "pageBreak");
      const divider = document.createElement("fluent-divider");
      divider.setAttribute("align-content", "center");
      divider.setAttribute("appearance", "subtle");
      divider.setAttribute("data-pb", "");
      divider.textContent = t("ribbon.cmd.page-break");
      dom.append(divider);
      return { dom, ignoreMutation: () => true };
    };
  },
});
