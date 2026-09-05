import type { ParagraphChild, RubyOptions } from "@office-open/docx";
import { Mark } from "@tiptap/core";

import { mergeTextNodes } from "../converters/styles";
import type { JSONContent } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/** `{ ruby }` is a RunOptions children shape (w:ruby lives INSIDE its w:r,
 *  like w:tab), not a ParagraphChild branch — so it has no Extract-able union
 *  member. The resolveRun children walk reaches the inline rules with items
 *  cast to ParagraphChild; the intersection keeps the predicate assignable to
 *  ParseInlineRule while letting convert read the real shape. */
type RubyBranch = ParagraphChild & { ruby: RubyOptions };

/** Concatenate the text of resolved inline content — the annotation text is
 *  stored flat on the mark (rt runs carry no formatting of their own beyond
 *  the ruby font size, which lives in the properties). */
function flattenText(nodes: JSONContent[]): string {
  let text = "";
  for (const node of nodes) {
    if (node.type === "text") text += node.text ?? "";
  }
  return text;
}

/** `{ ruby: {...} }` → text[] carrying a ruby mark. Mirrors resolveHyperlink:
 *  recurse the base runs via ctx, merge adjacent text, then stamp every text
 *  node with the mark. Returns null for an empty base. */
function resolveRuby(ruby: RubyOptions, ctx: ResolveContext): JSONContent[] | null {
  const props = ruby.properties;
  if (!props) return null;
  const base = ctx.resolveInlineChildren(ruby.base?.children ?? []);
  if (base.length === 0) return null;
  const attrs = {
    text: flattenText(ctx.resolveInlineChildren(ruby.text?.children ?? [])),
    alignment: props.alignment,
    fontSize: props.fontSize,
    baseFontSize: props.baseFontSize,
    raise: props.raise,
    languageId: props.languageId,
    dirty: props.dirty ?? null,
  };
  const merged = mergeTextNodes(base);
  for (const node of merged) {
    if (node.type === "text") {
      node.marks = [...(node.marks ?? []), { type: "ruby", attrs }];
    }
  }
  return merged;
}

// DOCX ruby run → office-open `{ ruby: {...} }` (the compile leg lives in
// DocxManager.compileTextRun next to the hyperlink container).
export const parseDocxInline: ParseInlineRule = {
  match: (child): child is RubyBranch => "ruby" in child,
  convert: (child, ctx) => resolveRuby((child as RubyBranch).ruby, ctx),
};

/**
 * Ruby — DOCX phonetic guide (拼音指南) conversion. The annotation rides a
 * character mark (attrs: annotation text + the CT_RubyPr fields) so the base
 * text stays directly editable and caret/selection mapping needs no special
 * casing; the w:ruby container shape is rebuilt on compile. `inclusive: false`
 * keeps typing at either edge outside the guide.
 */
export const Ruby = Mark.create({
  name: "ruby",
  inclusive: false,

  parseDocxInline,

  addAttributes() {
    return {
      // The annotation text (w:rt, flattened).
      text: { default: "" },
      // ST_RubyAlign token — how the annotation sits over the base text.
      alignment: { default: "center" },
      // Annotation font size in points (w:hps, halftwips/2 on parse).
      fontSize: { default: null },
      // Base-text font size in points (w:hpsBaseText).
      baseFontSize: { default: null },
      // Vertical offset of the annotation in points (w:hpsRaise).
      raise: { default: null },
      // Language identifier of the annotation (w:lid).
      languageId: { default: null },
      // Word's "recalculate phonetic guide" flag (w:dirty).
      dirty: { default: null },
    };
  },
});
