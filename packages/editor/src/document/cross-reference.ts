/**
 * Cross-reference (交叉引用) model shared by the dialog and the host commit:
 * the reference-type × content matrix (Word's Cross-reference dialog),
 * the REF/PAGEREF/NOTEREF instruction builder, the i18n key mapping, and the
 * candidate shape the document scan produces.
 *
 * Word's type list and its "Insert reference to" options:
 *
 * | type          | contents                                                |
 * |---------------|---------------------------------------------------------|
 * | Numbered item | paragraph number, page number, above/below              |
 * | Heading       | heading text, page number, heading number, above/below  |
 * | Bookmark      | bookmark text, page number, above/below                 |
 * | Footnote      | footnote number, page number, above/below               |
 * | Endnote       | endnote number, page number, above/below                |
 * | Figure/Table  | entire caption, label and number, caption text, …       |
 *
 * Equation is Word's seventh type; it defers to the equation workstream (G4)
 * and the dialog exposes it disabled.
 *
 * Resolution contract: references are inserted with a live-evaluated cache and
 * resolve at update time (F9 / Update All Fields) through the field engine's
 * bookmark table — the painter keeps the cached result, like DATE/REF in Word
 * until the user updates. `\p` (above/below) compares reading-order positions
 * rather than Word's same-page restriction.
 *
 * Deferred switches (not generated, rather than approximated):
 *  - REF `\r` / `\w` — the paragraph number in relative/full outline context.
 *    They need the outline-numbering state at every heading, which the runtime
 *    model only expands inside the layout projection; direct paragraph
 *    numbering is covered by `\n` (see the host scan's paragraph-number walk).
 *  - REF `\t` — suppressing a `\n` number's non-numeric text. The `\n` value
 *    is the level marker with trailing periods stripped, Word's numeric shape
 *    for the list formats the editor produces.
 *  - REF `\f` — a note's body text through a reference bookmark (NOTEREF
 *    carries the reference mark, which is what the dialog's number content
 *    needs).
 */

/** One reference type. Captions with any label share the `caption` kind; the
 *  dialog splits them by `label` (Figure/Table/…). */
export type CrossRefKind =
  | "numbered"
  | "heading"
  | "bookmark"
  | "footnote"
  | "endnote"
  | "caption"
  | "equation";

/** One "Insert reference to" choice, per type. */
export type CrossRefContent =
  | "text"
  | "label"
  | "entire"
  | "captionText"
  | "number"
  | "page"
  | "aboveBelow";

/** The type × content matrix in Word's dialog order. */
export const CROSS_REFERENCE_CONTENTS: Readonly<Record<CrossRefKind, readonly CrossRefContent[]>> =
  {
    numbered: ["number", "page", "aboveBelow"],
    heading: ["text", "page", "number", "aboveBelow"],
    bookmark: ["text", "page", "aboveBelow"],
    footnote: ["number", "page", "aboveBelow"],
    endnote: ["number", "page", "aboveBelow"],
    caption: ["entire", "label", "captionText", "page", "aboveBelow"],
    equation: [],
  };

/** The REF/PAGEREF/NOTEREF instruction for one target × content choice, or
 *  null when the combination is not part of the matrix. `\h` (hyperlink) rides
 *  every form — Word's dialog has "Insert as hyperlink" checked by default.
 *  Notes reference through NOTEREF (the reference mark); the other types
 *  through REF. */
export function crossReferenceInstruction(
  kind: CrossRefKind,
  content: CrossRefContent,
  name: string,
): string | null {
  if (!CROSS_REFERENCE_CONTENTS[kind].includes(content)) return null;
  const note = kind === "footnote" || kind === "endnote";
  if (content === "page") return `PAGEREF ${name} \\h`;
  if (content === "aboveBelow") return `${note ? "NOTEREF" : "REF"} ${name} \\p \\h`;
  if (note) return `NOTEREF ${name} \\h`;
  if (content === "number") return `REF ${name} \\n \\h`;
  return `REF ${name} \\h`;
}

/** The i18n key of one content choice's option label (the word follows the
 *  type, like Word: "Heading text" vs "Bookmark text", "Footnote number" vs
 *  "Paragraph number"). */
export function crossReferenceContentKey(kind: CrossRefKind, content: CrossRefContent): string {
  switch (content) {
    case "label":
      return "crossRef.refLabel";
    case "entire":
      return "crossRef.content.entire";
    case "captionText":
      return "crossRef.content.captionText";
    case "page":
      return "crossRef.refPage";
    case "aboveBelow":
      return "crossRef.content.aboveBelow";
    case "text":
      return kind === "heading" ? "crossRef.content.headingText" : "crossRef.refText";
    default:
      return kind === "heading"
        ? "crossRef.content.headingNumber"
        : kind === "footnote"
          ? "crossRef.content.footnoteNumber"
          : kind === "endnote"
            ? "crossRef.content.endnoteNumber"
            : "crossRef.content.number";
  }
}

/** The i18n key of one reference type's dialog label, or null when the label
 *  is rendered verbatim (a custom caption label like "Exhibit"). Caption
 *  labels share the caption dialog's words (`caption.figure`/`caption.table`),
 *  which are already the label nouns in both languages. */
export function crossReferenceTypeKey(kind: CrossRefKind, label?: string): string | null {
  switch (kind) {
    case "numbered":
      return "crossRef.type.numbered";
    case "heading":
      return "crossRef.type.heading";
    case "footnote":
      return "crossRef.type.footnote";
    case "endnote":
      return "crossRef.type.endnote";
    case "equation":
      return "caption.equation";
    case "caption": {
      const lower = (label ?? "").toLowerCase();
      if (lower === "figure") return "caption.figure";
      if (lower === "table") return "caption.table";
      if (lower === "equation") return "caption.equation";
      return null;
    }
    default:
      return "crossRef.bookmark";
  }
}

/** Whether a content choice has data to point at for this target. The dialog
 *  keeps unavailable options disabled and the commit refuses them — an
 *  unnumbered heading has no heading number, a caption without text no
 *  caption text; either would otherwise insert an empty reference. */
export function crossReferenceContentAvailable(
  target: CrossReferenceTarget,
  content: CrossRefContent,
): boolean {
  if (content === "number") return target.number != null;
  if (content === "captionText") return !!target.captionText;
  return true;
}

/** One referenceable target the document scan found. Paragraph targets carry
 *  the absolute ranges a missing hidden bookmark is created over; note
 *  targets anchor on their reference atom. */
export interface CrossReferenceTarget {
  /** Stable dialog/commit key (kind + identity). */
  key: string;
  kind: Exclude<CrossRefKind, "equation">;
  /** The target's existing bookmark name ("" when it has none yet — Word
   *  creates a hidden `_Ref…` bookmark on first reference). */
  name: string;
  /** The caption's label+number bookmark name (captions only) — the pair a
   *  "label and number" reference reuses. */
  labelName?: string;
  /** The caption's label word (Figure/Table/…); captions only. */
  label?: string;
  /** The anchor position: the paragraph's start, or the note reference atom. */
  pos: number;
  /** The referenced text: the bookmark's inner text, the heading/bookmark
   *  text, or the note's displayed number. */
  text: string;
  /** Whole-caption text, label + number + caption text (captions only). */
  captionFull?: string;
  /** The caption's text after the label and number (Word's "Only caption
   *  text"), the separator stripped. */
  captionText?: string;
  /** The paragraph's list/outline number (heading/numbered) or the note's
   *  displayed number (footnote/endnote), trailing periods stripped. */
  number?: string;
  /** Displayed page (1-based) when pagination is at hand. */
  page?: number;
  /** The dialog list's display string. */
  listText: string;
  /** Absolute range of the paragraph's inline content (the hidden bookmark a
   *  heading/numbered item gets); [pos, pos + 1) for note atoms. */
  contentFrom: number;
  contentTo: number;
  /** Absolute end of the caption's label+number (captions only) — where a
   *  missing `_Ref` bookmark is created. */
  labelTo?: number;
  /** Absolute position the caption text starts at (captions only, past the
   *  separator) — where an "Only caption text" bookmark starts. */
  captionTextFrom?: number;
  /** The existing target bookmark's inner range (when `name` is set). */
  bookmarkFrom?: number;
  bookmarkTo?: number;
}
