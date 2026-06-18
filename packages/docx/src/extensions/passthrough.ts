import { Node } from "../core";

/**
 * Passthrough — block atom carrying an opaque {@link SectionChild} that has
 * no native Tiptap representation (rawXml, bookmarkStart/End, toc, textbox,
 * altChunk, subDoc, customXml).
 *
 * The full SectionChild is stored as JSON in `attrs.data` so the DOCX→JSON→DOCX
 * round-trip stays byte-faithful: office-open's stringify handles the inner
 * structure verbatim (including a textbox's nested children, which remain as
 * structured ParagraphOptions inside the blob rather than editable Tiptap
 * nodes). The node is not editable — it renders a read-only placeholder in HTML.
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

  renderHTML({ node }: { node: { attrs: Record<string, unknown> } }) {
    let label = "DOCX";
    try {
      const parsed = JSON.parse((node.attrs.data as string) || "{}") as Record<string, unknown>;
      const key = Object.keys(parsed)[0];
      if (key) label = key;
    } catch {
      /* keep default label */
    }
    return [
      "div",
      {
        "data-passthrough": label,
        contenteditable: "false",
        style:
          "display:block;padding:0.5em 0.75em;margin:0.5em 0;border:1px dashed #bbb;border-radius:4px;color:#888;font-size:0.85em;background:#fafafa",
      },
      ["span", {}, `[${label}]`],
    ];
  },
});
