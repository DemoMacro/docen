import { Extension } from "@docen/docx/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { t } from "../../ui";

const SECTION_BREAK_KEY = new PluginKey("docen-section-break-marks");

/**
 * Editor-only rendering of section boundaries as a Fluent divider with a
 * centered label (Word's "Section Break" marker), mirroring the pageBreak
 * NodeView — shown only while formatting marks are on.
 *
 * A section break is NOT a node: OOXML sectPr rides on the section's last
 * paragraph's pPr, so the engine stamps `sectionProperties` on the paragraph
 * (see the SectionBreak command extension). With no node there is no NodeView,
 * so the marker is painted with a widget decoration placed right after each
 * section-carrying paragraph instead.
 *
 * The decoration is always inserted; visibility is driven by the host
 * [show-marks] attribute via CSS (same mechanism as page-break), so no toggle
 * state is needed — flipping show-marks shows/hides both markers at once.
 */
export const SectionBreakMarks = Extension.create({
  name: "sectionBreakMarks",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: SECTION_BREAK_KEY,
        props: {
          decorations(state) {
            const decos: Decoration[] = [];
            state.doc.descendants((node, pos) => {
              if (node.type.name !== "paragraph") return;
              if (node.attrs.sectionProperties == null) return;
              decos.push(
                Decoration.widget(pos + node.nodeSize, () => {
                  const wrap = document.createElement("div");
                  wrap.setAttribute("data-section-break", "");
                  wrap.contentEditable = "false";
                  const divider = document.createElement("fluent-divider");
                  divider.setAttribute("align-content", "center");
                  divider.setAttribute("appearance", "subtle");
                  divider.setAttribute("data-sb", "");
                  divider.textContent = t("ribbon.cmd.section-break");
                  wrap.append(divider);
                  return wrap;
                }),
              );
            });
            return DecorationSet.create(state.doc, decos);
          },
        },
      }),
    ];
  },
});
