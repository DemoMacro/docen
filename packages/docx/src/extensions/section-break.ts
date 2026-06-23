import { Extension } from "../core";

/**
 * SectionBreak — command extension that marks a paragraph as a section boundary.
 *
 * OOXML sections (sectPr) attach to a section's LAST paragraph's pPr, NOT a
 * standalone node. So this extension provides only the `setSectionBreak`
 * command (stamps sectionProperties on the current paragraph); the paragraph
 * extension carries the sectionProperties/sectionHeaders/sectionFooters attrs,
 * and DocxManager splits/merges sections in compile/resolve by reading them off
 * the paragraph.
 *
 * The final section's sectPr rides on doc.attrs.sectionProperties (it lives at
 * <w:body>'s end in OOXML). Single-section documents have no section-carrying
 * paragraph at all.
 */
export const SectionBreak = Extension.create({
  name: "sectionBreak",

  addCommands() {
    return {
      // Mark the current paragraph as a section's last paragraph. An empty
      // sectionProperties object closes a section on compile (office-open fills
      // A4/Normal defaults); callers can refine the geometry afterwards.
      setSectionBreak:
        () =>
        ({ commands }) =>
          commands.updateAttributes("paragraph", { sectionProperties: {} }),
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    sectionBreak: {
      /** Stamp sectionProperties on the current paragraph, ending a section. */
      setSectionBreak: () => ReturnType;
    };
  }
}
