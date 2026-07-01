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
    let key = "";
    try {
      const parsed = JSON.parse((node.attrs.data as string) || "{}") as Record<string, unknown>;
      key = Object.keys(parsed)[0] ?? "";
      if (key) label = key;
    } catch {
      /* keep default label */
    }
    // bookmarkStart/End are invisible position markers — a Word bookmark anchors
    // a range and has NO layout box. Render hidden so it occupies no space
    // (matching Word, where bookmarks are non-printing metadata); measure's
    // domHeightOf reads the hidden box as 0, so it takes no page space either.
    // Round-trip is unaffected: attrs.data still carries the SectionChild verbatim.
    if (key === "bookmarkStart" || key === "bookmarkEnd") {
      return [
        "div",
        { "data-passthrough": label, contenteditable: "false", style: "display:none" },
      ];
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

  renderHTML({ node }: { node: { attrs: Record<string, unknown> } }) {
    const data = (node.attrs.data as string) || "{}";
    let bookmarkName: string | undefined;
    // A bookmarkStart carries a name — expose it as the element id so anchor
    // links (#name, e.g. a TOC entry jumping to its heading) resolve to it.
    try {
      const parsed = JSON.parse(data) as { bookmarkStart?: { name?: string } };
      bookmarkName = parsed.bookmarkStart?.name;
    } catch {
      /* malformed data — no bookmark name */
    }
    const attrs: Record<string, string> = {
      "data-inline-passthrough": data,
      contenteditable: "false",
    };
    if (bookmarkName) {
      // Keep the anchor in-flow at zero size (not display:none) so it has a
      // layout box — ProseMirror's coordsAtPos/scrollIntoView can then resolve
      // the anchor position when a TOC link follows it. Other inline
      // passthrough children (bookmarkEnd, comment ranges, proofErr, …) carry
      // no anchor and stay display:none.
      attrs.id = bookmarkName;
      attrs.style = "display:inline-block;width:0;height:0;overflow:hidden;vertical-align:baseline";
    } else {
      attrs.style = "display:none";
    }
    return ["span", attrs];
  },
});
