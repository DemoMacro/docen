import { Node } from "../core";

/**
 * Passthrough — block atom carrying an opaque {@link SectionChild} that has
 * no native Tiptap representation (rawXml, bookmarkStart/End, textbox,
 * altChunk, subDoc, customXml).
 *
 * The full SectionChild is stored as JSON in `attrs.data` so the DOCX→JSON→DOCX
 * round-trip stays byte-faithful: office-open's stringify handles the inner
 * structure verbatim (including a textbox's nested children, which remain as
 * structured ParagraphOptions inside the blob rather than editable Tiptap
 * nodes). The node is not editable; the canvas paints its placeholder.
 *
 * DOCX serialization is inlined in DocxManager (compile/resolve read/write
 * `attrs.data` directly), so no renderDocx/parseDocx is needed here.
 */
export const Passthrough = Node.create({
  name: "passthrough",
  group: "block",
  atom: true,

  addAttributes() {
    return {
      data: {
        default: "{}",
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-passthrough]" }];
  },
});

/**
 * InlinePassthrough — inline atom carrying an opaque inline ParagraphChild that
 * has no native Tiptap representation (bookmarkStart/End, comment range markers,
 * proofErr, track-change markers, …). The full ParagraphChild rides in
 * `attrs.data` as JSON so DOCX→JSON→DOCX round-trips byte-faithful; the atom is
 * zero-width (bookmark/range markers carry no layout box), matching Word's
 * non-printing metadata. Mirrors the block-level Passthrough for inline children.
 */
export const InlinePassthrough = Node.create({
  name: "inlinePassthrough",
  group: "inline",
  inline: true,
  atom: true,

  addAttributes() {
    return {
      data: {
        default: "{}",
        rendered: false,
        parseHTML: (element: HTMLElement) =>
          element.getAttribute("data-inline-passthrough") ?? "{}",
      },
    };
  },

  parseHTML() {
    return [{ tag: "span[data-inline-passthrough]" }];
  },
});
