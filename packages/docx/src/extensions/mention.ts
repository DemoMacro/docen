import type { ParagraphChild, SdtRunOptions } from "@office-open/docx";
import { Mention as MentionBase } from "@tiptap/extension-mention";

import type { ParseInlineRule } from "./types";

/**
 * Mention extension — owns the DOCX expression of an inline mention.
 *
 * A mention is an atom node carrying `{ id, label }`. DOCX has no mention
 * element, but an inline text-SDT (CT_SdtRun) is a reversible carrier: the
 * `id` rides in the SDT alias, the `label` as the SDT's run text, and a fixed
 * tag marks the type so resolve can recover the mention. (customXml would also
 * carry the id but triggers Word's i4i patent warning; SDT does not.)
 */

/** SDT tag marking a mention content control. */
export const MENTION_TAG = "docen-mention";

/** Inline text-SDT carrying a mention (id in alias, label as run text). */
export function createMention(id: string, label: string): Record<string, unknown> {
  return {
    sdt: {
      properties: { tag: MENTION_TAG, alias: id, text: {} },
      children: [{ text: label }],
    },
  };
}

/** The inline-SDT ParagraphChild branch a mention rides in. */
type MentionBranch = Extract<ParagraphChild, { sdt: SdtRunOptions }>;

/** True if an inline SDT child carries a mention. */
export function isMention(child: ParagraphChild): boolean {
  return "sdt" in child && child.sdt.properties.tag === MENTION_TAG;
}

/** Read a mention SDT → `{ id, label }`. */
export function readMention(child: MentionBranch): { id: string; label: string } {
  const sdt = child.sdt;
  const id = sdt.properties.alias ?? "";
  let label = "";
  const first = sdt.children?.[0];
  if (typeof first === "string") label = first;
  else if (first && typeof first === "object" && "text" in first) label = String(first.text ?? "");
  return { id, label };
}

// DOCX inline text-SDT carrying a mention (CT_SdtRun) → office-open
// ParagraphChild `{ sdt: { properties: { tag } } }`. isMention guards the tag so
// a non-mention inline SDT falls through to the dispatcher's drop fallback.
export const parseDocxInline: ParseInlineRule<MentionBranch> = {
  match: (child): child is MentionBranch => isMention(child),
  convert: (child) => {
    const { id, label } = readMention(child);
    return { type: "mention", attrs: { id, label } };
  },
};

// DocxManager builds mention SDTs via createMention/readMention above; the
// extension itself carries no DOCX attrs, but declares the inline parse rule so
// resolve is reflective.
export const Mention = MentionBase.extend({ parseDocxInline });
