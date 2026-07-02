import { Mention as MentionBase } from "@tiptap/extension-mention";

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

/** True if an inline SDT child carries a mention. */
export function isMention(child: unknown): boolean {
  if (typeof child !== "object" || child === null || !("sdt" in child)) return false;
  const tag = (child as { sdt?: { properties?: { tag?: string } } }).sdt?.properties?.tag;
  return tag === MENTION_TAG;
}

/** Read a mention SDT → `{ id, label }`. */
export function readMention(child: unknown): { id: string; label: string } {
  const sdt = (child as { sdt?: { properties?: { alias?: string }; children?: unknown[] } }).sdt;
  const id = sdt?.properties?.alias ?? "";
  let label = "";
  const first = sdt?.children?.[0];
  if (typeof first === "string") label = first;
  else if (first && typeof first === "object" && "text" in first)
    label = String((first as { text?: string }).text ?? "");
  return { id, label };
}

// DocxManager builds/parses mention SDTs via createMention/readMention above;
// the extension itself carries no DOCX attrs of its own.
export { MentionBase as Mention };
