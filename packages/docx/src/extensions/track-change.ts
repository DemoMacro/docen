import type { ParagraphChild, RunOptions } from "@office-open/docx";

import { mergeTextNodes } from "../converters/styles";
import type { JSONContent } from "../core";
import { Mark } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

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
 * Container-level, NOT rPr-level: like `link`, these wrap child runs, so resolve
 * is declared via parseDocxInline (resolveTrackedChange) and compile via
 * compileTrackedChangeRun (compileTextRun pushes `{insertion|deletion:{...}}`).
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

// office-open ChangedProperties: { id:number; author:string; date:string }.
// `id` is a number on the OOXML side and is kept verbatim. All three are
// metadata — not rendered to HTML (the tag/class already identifies the mark).
const trackChangeAttrs = () => ({
  id: { default: null, rendered: false },
  author: { default: null, rendered: false },
  date: { default: null, rendered: false },
});

/** ParagraphChild `{ insertion|deletion: {...} }` → text[] carrying the mark.
 *  Mirrors the old DocxManager.resolveTrackedChange: recurse the container's
 *  runs via ctx, merge adjacent text, then stamp every text node with the
 *  revision mark alongside any existing rPr marks. Returns null for an empty
 *  container. */
function resolveTrackedChange(
  opts: {
    id?: number;
    author?: string;
    date?: string;
    children?: (RunOptions | string)[];
  },
  type: "insertion" | "deletion",
  ctx: ResolveContext,
): JSONContent[] | null {
  const content = ctx.resolveInlineChildren((opts.children ?? []).map((c) => c as ParagraphChild));
  if (content.length === 0) return null;
  const merged = mergeTextNodes(content);
  const mark = {
    type,
    attrs: {
      id: opts.id ?? null,
      author: opts.author ?? null,
      date: opts.date ?? null,
    },
  };
  for (const node of merged) {
    if (node.type === "text") {
      node.marks = [...(node.marks ?? []), mark];
    }
  }
  return merged;
}

// DOCX `<w:ins>` run → office-open ParagraphChild `{ insertion: {...} }`.
const insertionRule: ParseInlineRule = {
  match: (child) => "insertion" in child,
  convert: (child, ctx) =>
    resolveTrackedChange(
      (
        child as {
          insertion: {
            id?: number;
            author?: string;
            date?: string;
            children?: (RunOptions | string)[];
          };
        }
      ).insertion,
      "insertion",
      ctx,
    ),
};

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

  parseDocxInline: insertionRule,
});

// DOCX `<w:del>` run → office-open ParagraphChild `{ deletion: {...} }`.
const deletionRule: ParseInlineRule = {
  match: (child) => "deletion" in child,
  convert: (child, ctx) =>
    resolveTrackedChange(
      (
        child as {
          deletion: {
            id?: number;
            author?: string;
            date?: string;
            children?: (RunOptions | string)[];
          };
        }
      ).deletion,
      "deletion",
      ctx,
    ),
};

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

  parseDocxInline: deletionRule,
});
