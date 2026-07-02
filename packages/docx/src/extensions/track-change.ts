import { Mark } from "../core";

/**
 * Track Changes marks (Word revision tracking).
 *
 * OOXML records inline revisions as `<w:ins>` / `<w:del>` containers wrapping
 * runs (carrying w:author / w:date / w:id metadata). office-open models these
 * as `ParagraphChild.insertion` / `.deletion` — `{ id, author, date, children }`
 * — structurally identical to `hyperlink` (an inline container with attrs +
 * child runs). docen mirrors that as two Tiptap marks applied to the contained
 * text:
 *
 *  - `insertion` — text added by a reviewer (Word renders colored + underlined)
 *  - `deletion`  — text marked for removal (Word renders colored + strikethrough;
 *    the text stays visible until the change is accepted/rejected)
 *
 * Container-level, NOT rPr-level: like `link`, these wrap child runs, so they
 * are bridged directly in DocxManager (resolveParagraphChild →
 * resolveTrackedChange; compileTextRun → push `{insertion|deletion:{...}}`).
 * They do NOT use the renderDocx/parseDocx mark hook — that is for rPr-level
 * marks like strike/bold. The attrs (id/author/date) are round-tripped via
 * resolve/compile and kept out of HTML (`rendered:false`): HTML paste loses
 * the metadata but keeps the native `<ins>`/`<del>` tag (the class is a CSS
 * hook); DOCX round-trip is byte-faithful.
 *
 * HTML tags: `<ins>`/`<del>` are HTML's native editorial-revision elements, so
 * they are used instead of a bare span — semantic, accessible, and matching
 * browser defaults (underlined / struck-through). `<ins>` has no competing
 * mark, so both the classed tag and a bare pasted `<ins>` are claimed. `<del>`
 * is also matched by the base Strike mark, so only the classed tag is claimed
 * to avoid shadowing strike on a bare `<del>`.
 *
 * P1 scope: render + round-trip only. accept/reject commands, nested
 * revisions, block-level revisions, and format-revision (markChange) are out
 * of scope (office-open parses inline w:ins/w:del only).
 */

// office-open ChangedAttributesProperties: { id:number; author:string; date:string }.
// `id` is a number on the OOXML side and is kept verbatim. All three are
// metadata — not rendered to HTML (the tag/class already identifies the mark).
const trackChangeAttrs = () => ({
  id: { default: null, rendered: false },
  author: { default: null, rendered: false },
  date: { default: null, rendered: false },
});

export const Insertion = Mark.create({
  name: "insertion",
  // Read-only render for now: a caret inside a revision range must not extend
  // the mark onto newly typed text. Re-enable inclusivity once accept/reject
  // (P1.2) makes revision editing first-class.
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  renderHTML() {
    // <ins> is HTML's native "inserted text" element — semantic, accessible,
    // and browser-default underlined. The class stays as a CSS hook.
    return ["ins", { class: "docen-insertion" }, 0];
  },
  parseHTML() {
    // Claim our <ins class="docen-insertion"> plus a bare <ins> from pasted
    // HTML — no other mark matches <ins>, so the bare tag is safe.
    return [{ tag: "ins.docen-insertion" }, { tag: "ins" }];
  },
});

export const Deletion = Mark.create({
  name: "deletion",
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  renderHTML() {
    // <del> is HTML's native "deleted text" element — semantic and
    // browser-default struck-through.
    return ["del", { class: "docen-deletion" }, 0];
  },
  parseHTML() {
    // Only claim <del class="docen-deletion">: the base Strike mark already
    // matches a bare <del>, so claiming all <del> would shadow strike. DOCX
    // round-trip runs through resolve/compile, not this parseHTML.
    return [{ tag: "del.docen-deletion" }];
  },
});
