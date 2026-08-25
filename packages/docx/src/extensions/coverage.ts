import type { ParagraphChild, RunOptions, SectionChild } from "@office-open/docx";

/**
 * Concept-coverage disposition tables — the explicit registry of how EVERY
 * office-open document-model branch is handled by the Tiptap layer.
 *
 * Two directions are covered:
 *  - resolve (office-open → Tiptap): every SectionChild / ParagraphChild branch
 *    is claimed by an editable node, carried verbatim through a passthrough
 *    atom, or dropped for a logged reason.
 *  - compile (Tiptap → office-open): editor-only nodes that do NOT map back.
 *    Known losses, registered here so they are decisions, not accidents:
 *      • `mathematics` — KaTeX latex has no OMML bridge (an office-open
 *        MathInput tree is not derivable from latex); compile silently skips it.
 *      • `codeBlock.language` — the Code style has no language slot.
 *
 * The satisfies guards make the registry complete by construction: a new
 * office-open union branch widens the tag type, a missing entry fails the
 * build; an entry naming no real branch also fails. This file is consumed by
 * coverage.spec.ts, which drives round-trip fixtures per tag — the registry
 * claims, the spec proves.
 */

/** Extract the tag key of every single-key branch member of a union. MUST be
 *  distributive (naked `U extends object`): a bare `keyof UnionType` would be
 *  the INTERSECTION of member keys — never, for tag unions with no shared key —
 *  and Record<never, …> would satisfy any table (the guard would silently pass).
 *  Branches carrying optional companion keys (e.g. {footnoteReference,
 *  properties?}) yield those too — filtered by the caller's NonTag set. */
type TagOf<U> = U extends object ? { [K in keyof U]-?: K }[keyof U] : never;

/** Keys that appear on union branches but are NOT tags: the companion optional
 *  keys plus everything the untagged RunOptions fallback member contributes.
 *  keyof RunOptions stays a live reference, so a new RunOptions field never
 *  false-flags as a missing tag. Exceptions — office-open dual-models these as
 *  BOTH a run property and a real ParagraphChild branch; the branch wins as a
 *  tag: math (w:rPr flag vs {math: MathInput[]}), footnoteReference and
 *  endnoteReference (run-level w:*Reference vs the reference branch). */
type NonTagKeys = Exclude<
  "properties" | keyof RunOptions,
  "math" | "footnoteReference" | "endnoteReference"
>;

type ParagraphChildTag = Exclude<TagOf<ParagraphChild>, NonTagKeys>;
type SectionChildTag = TagOf<SectionChild>;

/** How a branch is handled: claimed by an editable extension (value names the
 *  claiming route), carried verbatim by a passthrough atom, or dropped (with
 *  the reason — a drop must always be a decision). */
export type Disposition =
  | { editable: string }
  | { passthrough: string }
  | { dropped: { reason: string } };

/** SectionChild branches (block level) — resolve direction. */
export const SECTION_CHILD_DISPOSITIONS = {
  paragraph: { editable: "Paragraph/Heading nodes" },
  table: { editable: "Table.parseDocxBlock" },
  toc: { editable: "tocField node (entries editable, field switches opaque)" },
  textbox: { passthrough: "block Passthrough atom" },
  sdt: { passthrough: "block Passthrough atom" },
  altChunk: { passthrough: "block Passthrough atom" },
  customXml: { passthrough: "block Passthrough atom" },
  bookmarkStart: { passthrough: "block Passthrough atom" },
  bookmarkEnd: { passthrough: "block Passthrough atom" },
  rawXml: { passthrough: "block Passthrough atom" },
} satisfies Record<SectionChildTag, Disposition>;

/** ParagraphChild branches (inline level) — resolve direction. Unclaimed
 *  shapes all land in the inlinePassthrough atom (verbatim, byte-faithful).
 *  The untagged RunOptions fallback member is editable by construction: it
 *  resolves to text nodes + marks (see resolveRun in converters/docx.ts) —
 *  its keys are all NonTag, so it has no row here. */
export const PARAGRAPH_CHILD_DISPOSITIONS = {
  picture: { editable: "image node" },
  hyperlink: { editable: "Link mark (container)" },
  insertion: { editable: "Insertion mark (container)" },
  deletion: { editable: "Deletion mark (container)" },
  pageBreak: { editable: "pageBreak node" },
  columnBreak: { editable: "columnBreak node" },
  wpsShape: { editable: "wpsShape node (text body editable, geometry opaque)" },
  wpgGroup: { editable: "wpgGroup node (whole tree opaque)" },
  bookmarkStart: { passthrough: "inlinePassthrough atom (name exposed for TOC anchors)" },
  bookmarkEnd: { passthrough: "inlinePassthrough atom" },
  bookmark: { passthrough: "inlinePassthrough atom" },
  chart: { passthrough: "inlinePassthrough atom" },
  smartArt: { passthrough: "inlinePassthrough atom" },
  math: { passthrough: "inlinePassthrough atom" },
  symbolRun: { passthrough: "inlinePassthrough atom" },
  footnoteReference: { passthrough: "inlinePassthrough atom" },
  endnoteReference: { passthrough: "inlinePassthrough atom" },
  commentRangeStart: { passthrough: "inlinePassthrough atom" },
  commentRangeEnd: { passthrough: "inlinePassthrough atom" },
  commentReference: { passthrough: "inlinePassthrough atom" },
  comment: { passthrough: "inlinePassthrough atom" },
  object: { passthrough: "inlinePassthrough atom" },
  pict: { passthrough: "inlinePassthrough atom" },
  contentPart: { passthrough: "inlinePassthrough atom" },
  proofErr: { passthrough: "inlinePassthrough atom" },
  positionalTab: { passthrough: "inlinePassthrough atom" },
  permStart: { passthrough: "inlinePassthrough atom" },
  permEnd: { passthrough: "inlinePassthrough atom" },
  moveFromRangeStart: { passthrough: "inlinePassthrough atom" },
  moveFromRangeEnd: { passthrough: "inlinePassthrough atom" },
  moveToRangeStart: { passthrough: "inlinePassthrough atom" },
  moveToRangeEnd: { passthrough: "inlinePassthrough atom" },
  movedFrom: { passthrough: "inlinePassthrough atom" },
  movedTo: { passthrough: "inlinePassthrough atom" },
  moveFrom: { passthrough: "inlinePassthrough atom" },
  moveTo: { passthrough: "inlinePassthrough atom" },
  customXmlInsRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlInsRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlDelRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlDelRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlMoveFromRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlMoveFromRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlMoveToRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlMoveToRangeEnd: { passthrough: "inlinePassthrough atom" },
  simpleField: { passthrough: "inlinePassthrough atom" },
  formField: { passthrough: "inlinePassthrough atom" },
  complexField: { passthrough: "inlinePassthrough atom" },
  seqIdentifier: { passthrough: "inlinePassthrough atom" },
  pageReference: { passthrough: "inlinePassthrough atom" },
  dir: { passthrough: "inlinePassthrough atom" },
  bdo: { passthrough: "inlinePassthrough atom" },
  smartTag: { passthrough: "inlinePassthrough atom" },
  customXml: { passthrough: "inlinePassthrough atom" },
  sdt: { passthrough: "inlinePassthrough atom" },
  subDoc: { passthrough: "inlinePassthrough atom" },
  rawXml: { passthrough: "inlinePassthrough atom" },
} satisfies Record<ParagraphChildTag, Disposition>;

/** RunOptions children shapes the resolve side drops (the children walk in
 *  resolveRun, converters/docx.ts). Not keyed by a union — an explicit
 *  decision list of the shapes with NO owning inline rule and no `break`:
 *  rule-owned children (tab, pageBreak, picture, …) are handled; a nested
 *  ParagraphChild member with no owning rule (e.g. {object}) drops there too
 *  but keeps its top-level disposition above — office-open parse emits those
 *  top-level, never nested. */
export const RUN_CHILDREN_DROPPED: readonly { tag: string; reason: string }[] = [
  { tag: "lastRenderedPageBreak", reason: "renderer pagination hint, not content" },
  { tag: "noBreakHyphen", reason: "no editable equivalent (would corrupt words)" },
  { tag: "softHyphen", reason: "no editable equivalent (would corrupt words)" },
  { tag: "carriageReturn", reason: "vertical-tab break; only w:br is modeled" },
  { tag: "separator", reason: "footnote/endnote separator run, not body content" },
  { tag: "continuationSeparator", reason: "footnote/endnote separator run, not body content" },
  { tag: "annotationRef", reason: "comment anchor marker, not body content" },
  { tag: "footnoteRef", reason: "footnote anchor marker, not body content" },
  { tag: "endnoteRef", reason: "endnote anchor marker, not body content" },
  { tag: "pgNum", reason: "live field, value recomputed at view time" },
  { tag: "dayShort", reason: "live date field, value recomputed at view time" },
  { tag: "dayLong", reason: "live date field, value recomputed at view time" },
  { tag: "monthShort", reason: "live date field, value recomputed at view time" },
  { tag: "monthLong", reason: "live date field, value recomputed at view time" },
  { tag: "yearShort", reason: "live date field, value recomputed at view time" },
  { tag: "yearLong", reason: "live date field, value recomputed at view time" },
  { tag: "ruby", reason: "phonetic guide has no editable representation" },
];
