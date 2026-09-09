import type { ParagraphChild, SectionChild } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

import { cleanAttrs } from "../converters/styles";
import { Node } from "../core";
import type { ParseBlockRule, ParseInlineRule } from "./types";
import { attrNative } from "./utils";

/**
 * Structured document tags (w:sdt) — content controls whose CONTENT is
 * editable while the control's own settings ride verbatim on attrs
 * (`properties` = the whole CT_SdtPr: tag/alias/checkbox/date/dropdown/…;
 * `endProperties` = w:sdtEndPr). Word's default view shows the content bare
 * with no chrome; the editor does the same, so nothing is lost visually —
 * the boundary box / title-bar control UI is a later concern, not a
 * fidelity one.
 *
 * Two nodes mirror the two OOXML shapes: the block control (CT_SdtBlock)
 * wraps a SectionChild[] stream like tocField does; the inline control
 * (CT_SdtRun) wraps a ParagraphChild[] run stream.
 */

type SdtBlockBranch = Extract<SectionChild, { sdt: unknown }>;
type SdtInlineBranch = Extract<ParagraphChild, { sdt: unknown }>;

/** Shared attrs assembly: control settings verbatim, undefined keys omitted. */
function sdtAttrs(sdt: {
  properties?: unknown;
  endProperties?: unknown;
}): Record<string, unknown> | undefined {
  const attrs = cleanAttrs({
    properties: sdt.properties,
    endProperties: sdt.endProperties,
  } as Record<string, unknown>);
  return Object.keys(attrs).length > 0 ? attrs : undefined;
}

// ── Block control (CT_SdtBlock) ──

export const parseDocxBlock: ParseBlockRule<SdtBlockBranch> = {
  match: (child): child is SdtBlockBranch => "sdt" in child,
  convert: (child, ctx) => {
    const content = ctx.resolveBlockStream((child.sdt.children ?? []) as SectionChild[]);
    if (content.length === 0) content.push({ type: "paragraph" });
    const node: JSONContent = { type: "sdtBlock", content };
    const attrs = sdtAttrs(child.sdt);
    if (attrs) node.attrs = attrs;
    return node;
  },
};

export const SdtBlock = Node.create({
  name: "sdtBlock",
  group: "block",
  content: "block+",
  // Content-control boundaries: isolating keeps Backspace at the start from
  // pulling the first inner paragraph out of the control; defining keeps the
  // shell when the whole content is selected and replaced.
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      properties: attrNative(),
      endProperties: attrNative(),
    };
  },

  parseHTML() {
    return [{ tag: "div.docx-sdt-block" }];
  },

  parseDocxBlock,
});

// ── Inline control (CT_SdtRun) ──

export const parseDocxInline: ParseInlineRule<SdtInlineBranch> = {
  match: (child): child is SdtInlineBranch => "sdt" in child,
  convert: (child, ctx) => {
    const content = ctx.resolveInlineChildren(
      (child.sdt.children ?? []) as (ParagraphChild | string)[],
    );
    const node: JSONContent = {
      type: "sdtInline",
      // inline+ requires content; an empty control carries one empty text node.
      content: content.length > 0 ? content : [{ type: "text", text: "" }],
    };
    const attrs = sdtAttrs(child.sdt);
    if (attrs) node.attrs = attrs;
    return node;
  },
};

export const SdtInline = Node.create({
  name: "sdtInline",
  group: "inline",
  inline: true,
  content: "inline+",
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      properties: attrNative(),
      endProperties: attrNative(),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-sdt-inline" }];
  },

  parseDocxInline,
});
