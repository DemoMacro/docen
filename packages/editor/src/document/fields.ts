/**
 * Field catalog + evaluator registry — the insert/update face of Word's field
 * dialog (插入 → 域). Fields ride the `inlinePassthrough` atom (`attrs.data`
 * carrying the office-open `simpleField`/`complexField`/`formField` branch
 * verbatim), so everything here is plain JSON in, plain JSON out; nothing
 * touches the PM doc.
 *
 * The catalog lists the insertable fields; {@link FIELD_EVALUATORS} re-derives
 * any field's value from a {@link FieldContext} — the edit-time update
 * commands (F9 / Update All Fields) and the render pass's live numbering
 * fields both go through {@link resolveField}. A null result means the context
 * cannot provide the value; the caller keeps the cached result, Word's
 * unresolved-field behavior (cached, else empty).
 */

import { formatNumber } from "@docen/layout";

/** One insertable field: the OOXML name and its default instruction (what the
 *  dialog's field-code box prefills). Field names stay English — Word's field
 *  dialog shows the field codes verbatim in every UI language. */
export interface FieldDef {
  name: string;
  instruction: string;
}

export interface FieldCategory {
  /** Suffix under `field.category.` */
  key: string;
  fields: FieldDef[];
}

/** Date fields share the "update automatically" shape the Date and Time
 *  dialog already inserts (`DATE \@ "…"`, see #insertDateTime). */
const date = (name: string, def: string): FieldDef => ({
  name,
  instruction: `${name} \\@ "${def}"`,
});

/** Fields whose instruction is the bare name (no picture, no switches). */
const plain = (name: string): FieldDef => ({ name, instruction: name });

export const FIELD_CATEGORIES: readonly FieldCategory[] = [
  {
    key: "date",
    fields: [
      date("DATE", "yyyy/M/d"),
      date("TIME", "H:mm:ss"),
      date("CREATEDATE", "yyyy/M/d H:mm"),
      date("SAVEDATE", "yyyy/M/d H:mm"),
      date("PRINTDATE", "yyyy/M/d"),
    ],
  },
  {
    key: "document",
    fields: [
      plain("AUTHOR"),
      plain("TITLE"),
      plain("SUBJECT"),
      plain("KEYWORDS"),
      plain("COMMENTS"),
      plain("FILENAME"),
      plain("REVNUM"),
      plain("NUMCHARS"),
      plain("NUMWORDS"),
    ],
  },
  {
    key: "numbering",
    fields: [
      // Dynamic in the painter (resolved per page) — the cached value is just
      // the measuring placeholder Word also caches. SECTION/SECTIONPAGES are
      // re-resolved by the render pass against the current pagination.
      plain("PAGE"),
      plain("NUMPAGES"),
      plain("SECTION"),
      plain("SECTIONPAGES"),
    ],
  },
];

/** A field atom under the caret, resolved from its passthrough branch — the
 *  instruction, the cached result, and a checkbox's checked state. */
export interface FieldRef {
  kind: "simpleField" | "complexField" | "formField";
  instruction?: string;
  result?: string;
  /** formField checkBox only. */
  checked?: boolean;
}

/** The field a passthrough branch carries, or null (not a field). Both the
 *  flat shapes office-open parses (instruction/cachedValue, instruction/result)
 *  and the checkbox's `checked` flag are read. */
export function fieldRef(branch: Record<string, unknown>): FieldRef | null {
  const simple = branch.simpleField;
  if (simple && typeof simple === "object") {
    const s = simple as { instruction?: unknown; cachedValue?: unknown };
    return {
      kind: "simpleField",
      instruction: typeof s.instruction === "string" ? s.instruction : undefined,
      result: typeof s.cachedValue === "string" ? s.cachedValue : undefined,
    };
  }
  const complex = branch.complexField;
  if (complex && typeof complex === "object") {
    const c = complex as { instruction?: unknown; result?: unknown };
    return {
      kind: "complexField",
      instruction: typeof c.instruction === "string" ? c.instruction : undefined,
      result: typeof c.result === "string" ? c.result : undefined,
    };
  }
  const form = branch.formField;
  if (form && typeof form === "object") {
    const box = (form as { checkBox?: unknown }).checkBox;
    const checked =
      box && typeof box === "object" ? (box as { checked?: unknown }).checked : undefined;
    return { kind: "formField", checked: checked === true };
  }
  return null;
}

/** What 更新域 needs to re-derive a value. Every part is optional so the same
 *  evaluator registry serves the edit-time update commands (document state
 *  only) and the render pass (page frame included). */
export interface FieldContext {
  /** "Now" for DATE/TIME (injected — evaluation must stay testable). The
   *  render pass creates one Date per render so a single render is stable. */
  now?: Date;
  /** The doc's core properties (creator/title/subject/keywords/comments/
   *  created/modified/lastPrinted/revision) backing the document-information
   *  fields. */
  core?: Record<string, unknown>;
  /** The document's word and character totals (Word's NUMWORDS/NUMCHARS). */
  words?: number;
  chars?: number;
  /** The open document's name, backing FILENAME. */
  filename?: string;
  /** The revision number (REVNUM) — core.xml cp:revision. */
  revision?: number;
  /** Custom document properties (DOCPROPERTY), name → value. */
  customProperties?: ReadonlyMap<string, string>;
  /** REF/PAGEREF/NOTEREF targets: bookmark name → its inner text, list
   *  number, start page, and reading-order position. */
  bookmarks?: ReadonlyMap<string, FieldBookmark>;
  /** The field atom's own document position — the `\p` switch's above/below
   *  compares it with the target bookmark's position. Absent = no `\p`. */
  selfPos?: number;
  /** Localized "above"/"below" for the `\p` switch (Word renders the UI
   *  language's words). Absent = English. */
  positionTerms?: { above: string; below: string };
  /** SEQ label → the ordinal this occurrence takes; the update walk assigns
   *  them in document order (each label restarts at 1, and at a heading the
   *  field's `\s` level marks). */
  sequences?: ReadonlyMap<string, number>;
  /** SEQ `\s <level>` chapter numbers in effect at the field's position:
   *  heading level (1-9) → the chapter number text (Word's "Include chapter
   *  number"). Absent = no heading of that level precedes the field. */
  chapters?: ReadonlyMap<number, string>;
  /** Caption label → the character between the chapter number and the
   *  sequence number (w:caption@w:sep, Word's "Use separator"). A label
   *  without an entry uses the hyphen default. */
  captionSeparators?: ReadonlyMap<string, string>;
  /** The rendered page frame (PAGE/NUMPAGES/SECTION/SECTIONPAGES/PAGEREF).
   *  Absent = no pagination at hand; page-dependent fields keep their cache. */
  frame?: FieldFrame;
}

/** One page slice of the pagination result, as the field evaluators read it. */
export interface FieldFrame {
  /** 1-based page number as shown (the section's w:pgNumType restart applied). */
  page?: number;
  /** Total pages in the document. */
  pageCount?: number;
  /** 1-based section number. */
  section?: number;
  /** Pages in the field's section. */
  sectionPages?: number;
  /** The section's w:numFmt token (absent = decimal). */
  pageFormat?: string;
}

/** One REF/PAGEREF/NOTEREF target: the bookmark's inner text (REF/NOTEREF),
 *  its paragraph's list/outline number (REF `\n`), the displayed page its
 *  bookmark start sits on (PAGEREF), and its reading-order position (`\p`).
 *  `page` is the section-numbering number (restart applied), `pageFormat` the
 *  target section's w:numFmt token — a bookmark crossing into another section
 *  formats with its own. */
export interface FieldBookmark {
  text?: string;
  /** The bookmarked paragraph's list number with trailing periods stripped
   *  (Word's REF `\n` shape) — absent when the paragraph carries no numbering
   *  the runtime can compute. */
  number?: string;
  page?: number;
  pageFormat?: string;
  /** The target's reading-order document position — the `\p` switch's
   *  above/below comparison. */
  pos?: number;
}

/** A parsed instruction: the field name plus its positional arguments and
 *  switches (`\* MERGEFORMAT`, `\@ "yyyy/M/d"`, `\h`). Quoted strings stay
 *  whole; a switch consumes the next token as its value. */
export interface ParsedFieldInstruction {
  /** The uppercase field name (first token). */
  name: string;
  /** The instruction verbatim, trimmed. */
  raw: string;
  /** Positional arguments after the name (`REF _Ref1`, `SEQ 图` → ["_Ref1"]),
   *  switch tokens and their values excluded. */
  args: string[];
  /** Switch values keyed without the backslash: `*` → "MERGEFORMAT",
   *  `@` → "yyyy/M/d", `#` → "0", `h` → "" (a flag switch). */
  switches: Record<string, string>;
}

export function parseFieldInstruction(instruction: string): ParsedFieldInstruction {
  const raw = instruction.trim();
  const tokens: string[] = [];
  let current = "";
  let quoted = false;
  for (const ch of raw) {
    if (ch === '"') {
      quoted = !quoted;
      continue;
    }
    if (!quoted && (ch === " " || ch === "\t")) {
      if (current) {
        tokens.push(current);
        current = "";
      }
      continue;
    }
    current += ch;
  }
  if (current) tokens.push(current);
  const name = (tokens[0] ?? "").toUpperCase();
  const args: string[] = [];
  const switches: Record<string, string> = {};
  for (let i = 1; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (!token.startsWith("\\") || token.length < 2) {
      args.push(token);
      continue;
    }
    const key = token[1]!;
    const next = tokens[i + 1];
    if (next && !next.startsWith("\\")) {
      switches[key] = next;
      i += 1;
    } else {
      switches[key] = "";
    }
  }
  return { name, raw, args, switches };
}

/** Formats `d` per a Word date picture (`\@ "yyyy年M月d日"`): the common
 *  tokens longest-first; an unrecognized picture falls back to the unformatted
 *  locale string, never an empty result. */
export function formatDate(d: Date, picture: string): string {
  const pad = (n: number): string => String(n).padStart(2, "0");
  const tokens: [RegExp, string][] = [
    [/yyyy/, String(d.getFullYear())],
    [/yy/, String(d.getFullYear() % 100).padStart(2, "0")],
    [/MMMM/, d.toLocaleString("zh-CN", { month: "long" })],
    [/MM/, pad(d.getMonth() + 1)],
    [/M/, String(d.getMonth() + 1)],
    [/dddd/, d.toLocaleString("zh-CN", { weekday: "long" })],
    [/ddd/, d.toLocaleString("zh-CN", { weekday: "short" })],
    [/dd/, pad(d.getDate())],
    [/d/, String(d.getDate())],
    [/HH/, pad(d.getHours())],
    [/H/, String(d.getHours())],
    [/hh/, pad(d.getHours() % 12 || 12)],
    [/h/, String(d.getHours() % 12 || 12)],
    [/mm/, pad(d.getMinutes())],
    [/m/, String(d.getMinutes())],
    [/ss/, pad(d.getSeconds())],
    [/s/, String(d.getSeconds())],
    [/AM\/PM/, d.getHours() < 12 ? "AM" : "PM"],
  ];
  let out = "";
  let i = 0;
  outer: while (i < picture.length) {
    for (const [re, value] of tokens) {
      const rest = picture.slice(i);
      const m = re.exec(rest);
      if (m && m.index === 0) {
        out += value;
        i += m[0].length;
        continue outer;
      }
    }
    out += picture[i];
    i += 1;
  }
  return out;
}

/** The shared string coercion for core-property values (a serialized core
 *  object may carry numbers/strings; anything else is absent). */
const str = (v: unknown): string | null =>
  typeof v === "string" ? v : typeof v === "number" ? String(v) : null;

/** A numeric context value as a finite number, or null when absent/malformed
 *  (a garbage revision must not paint "NaN"). Strings must parse to a finite
 *  number; booleans/null/objects/NaN/Infinity are absent. */
export function finiteNumber(v: unknown): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string" && v.trim() !== "") {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/** A core-property date as a Date, or null when absent/malformed. */
function coreDate(v: unknown): Date | null {
  if (v instanceof Date) return v;
  if (typeof v === "string" || typeof v === "number") {
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/** Date/time pictures — Word's defaults for the bare field names. */
const DATE_PICTURES: Record<string, string> = {
  DATE: "yyyy/M/d",
  TIME: "H:mm:ss",
  CREATEDATE: "yyyy/M/d H:mm",
  SAVEDATE: "yyyy/M/d H:mm",
  PRINTDATE: "yyyy/M/d",
};

const dateEvaluator =
  (key: "created" | "modified" | "lastPrinted", fallback: string): FieldEvaluator =>
  (field, ctx) => {
    const d = coreDate(ctx.core?.[key]);
    return d ? formatDate(d, field.switches["@"] ?? fallback) : null;
  };

/** DOCPROPERTY name → value from a `documentExtras` slice's round-tripped
 *  office-open `{ name, value }` entries; undefined when none. */
export function customPropertiesOf(
  extras: Record<string, unknown> | undefined,
): ReadonlyMap<string, string> | undefined {
  const raw = extras?.customProperties;
  if (!Array.isArray(raw)) return undefined;
  const map = new Map<string, string>();
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") continue;
    const { name, value } = entry as { name?: unknown; value?: unknown };
    if (typeof name !== "string" || value == null) continue;
    map.set(name, variantText(value));
  }
  return map.size > 0 ? map : undefined;
}

/** A custom-property variant as its display text (string/number/boolean pass,
 *  dates ISO-format; anything else is its JSON). */
function variantText(value: unknown): string {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean")
    return String(value);
  if (value instanceof Date) return value.toISOString();
  return JSON.stringify(value);
}

/** Fields whose value is a pure function of the live pagination — the render
 *  pass re-resolves these after every layout (Word keeps page numbers live
 *  through pagination). Everything else is update-time: F9 / Update All
 *  Fields / w:updateFields on open keeps the last result, Word's behavior. */
export const LIVE_FIELD_NAMES: ReadonlySet<string> = new Set([
  "PAGE",
  "NUMPAGES",
  "SECTION",
  "SECTIONPAGES",
]);

/** INFO <argument> aliases — the document-information names Word's INFO
 *  field forwards to the same evaluators. */
const INFO_ALIASES: Record<string, string> = {
  author: "AUTHOR",
  title: "TITLE",
  subject: "SUBJECT",
  keywords: "KEYWORDS",
  comments: "COMMENTS",
  createdate: "CREATEDATE",
  savedate: "SAVEDATE",
  printdate: "PRINTDATE",
  filename: "FILENAME",
  numpages: "NUMPAGES",
  numwords: "NUMWORDS",
  numchars: "NUMCHARS",
  revnum: "REVNUM",
  date: "DATE",
  time: "TIME",
};

/** The evaluator registry, keyed by field name. A null result means "this
 *  field has no value in this context" — the caller keeps the cached result
 *  (Word's unresolved-field behavior is the cached value, else empty). */
export type FieldEvaluator = (field: ParsedFieldInstruction, ctx: FieldContext) => string | null;

/** SEQ `\*` switch token → the number-format token {@link formatNumber}
 *  renders, keyed by the canonical tokens the Insert Caption dialog writes
 *  (the settings' w:numFmt values). Word's default (also the fallback for an
 *  unknown switch) is ARABIC. */
export const SEQ_NUMBER_FORMATS: Readonly<Record<string, string>> = {
  ARABIC: "decimal",
  ROMAN: "upperRoman",
  roman: "lowerRoman",
  ALPHABETIC: "upperLetter",
  alphabetic: "lowerLetter",
};

/** Render one sequence ordinal under its `\*` switch — the caption number
 *  (Word's Insert Caption format list). Word reads the switch word
 *  case-insensitively and its casing picks the output case: Roman/ROMAN → I,
 *  roman → i (ECMA's canonical uppercase is `Roman`); Alphabetic/ALPHABETIC →
 *  A, alphabetic → a; Arabic is case-free decimal. An unknown switch falls
 *  back to ARABIC. */
export function formatSeqNumber(switchToken: string | undefined, ordinal: number): string {
  const token = switchToken ?? "";
  switch (token.toLowerCase()) {
    case "roman":
      return formatNumber(token === "roman" ? "lowerRoman" : "upperRoman", ordinal);
    case "alphabetic":
      return formatNumber(token === "alphabetic" ? "lowerLetter" : "upperLetter", ordinal);
    default:
      return formatNumber("decimal", ordinal);
  }
}

/** The `\s` switch as a valid chapter level: an integer 1-9, or undefined —
 *  Word ignores any other switch value. */
export function seqChapterLevel(switchValue: string | undefined): number | undefined {
  const level = finiteNumber(switchValue);
  return level != null && Number.isInteger(level) && level >= 1 && level <= 9 ? level : undefined;
}

/** w:caption@w:sep token → the character between a chapter number and a SEQ
 *  number (Word's "Use separator" list). */
export const CAPTION_SEPARATOR_CHARS: Readonly<Record<string, string>> = {
  hyphen: "-",
  period: ".",
  colon: ":",
  emDash: "\u2014",
  enDash: "\u2013",
};

/** The `\p` switch's above/below — "above" when the target bookmark sits
 *  before the field in reading order, "below" otherwise. Word restricts the
 *  switch to a target on the field's own page; docen compares document
 *  positions, the order the pages themselves flow in. */
function aboveBelow(ctx: FieldContext, bookmark: FieldBookmark): string | null {
  if (bookmark.pos == null || ctx.selfPos == null) return null;
  const terms = ctx.positionTerms ?? { above: "above", below: "below" };
  return bookmark.pos < ctx.selfPos ? terms.above : terms.below;
}

export const FIELD_EVALUATORS: Readonly<Record<string, FieldEvaluator>> = {
  PAGE: (_field, ctx) =>
    ctx.frame?.page != null ? formatNumber(ctx.frame.pageFormat, ctx.frame.page) : null,
  NUMPAGES: (_field, ctx) => (ctx.frame?.pageCount != null ? String(ctx.frame.pageCount) : null),
  SECTION: (_field, ctx) =>
    ctx.frame?.section != null ? formatNumber(ctx.frame.pageFormat, ctx.frame.section) : null,
  SECTIONPAGES: (_field, ctx) =>
    ctx.frame?.sectionPages != null
      ? formatNumber(ctx.frame.pageFormat, ctx.frame.sectionPages)
      : null,
  DATE: (field, ctx) =>
    formatDate(ctx.now ?? new Date(), field.switches["@"] ?? DATE_PICTURES.DATE!),
  TIME: (field, ctx) =>
    formatDate(ctx.now ?? new Date(), field.switches["@"] ?? DATE_PICTURES.TIME!),
  CREATEDATE: dateEvaluator("created", DATE_PICTURES.CREATEDATE!),
  SAVEDATE: dateEvaluator("modified", DATE_PICTURES.SAVEDATE!),
  PRINTDATE: dateEvaluator("lastPrinted", DATE_PICTURES.PRINTDATE!),
  AUTHOR: (_field, ctx) => str(ctx.core?.creator),
  TITLE: (_field, ctx) => str(ctx.core?.title),
  SUBJECT: (_field, ctx) => str(ctx.core?.subject),
  KEYWORDS: (_field, ctx) => str(ctx.core?.keywords),
  COMMENTS: (_field, ctx) => str(ctx.core?.description),
  FILENAME: (_field, ctx) => ctx.filename ?? null,
  REVNUM: (_field, ctx) => {
    const revision = finiteNumber(ctx.revision) ?? finiteNumber(ctx.core?.revision);
    return revision != null ? String(revision) : null;
  },
  NUMWORDS: (_field, ctx) => (ctx.words != null ? String(ctx.words) : null),
  NUMCHARS: (_field, ctx) => (ctx.chars != null ? String(ctx.chars) : null),
  REF: (field, ctx) => {
    const bookmark = field.args[0] ? ctx.bookmarks?.get(field.args[0]) : undefined;
    if (!bookmark) return null;
    if ("p" in field.switches) return aboveBelow(ctx, bookmark);
    if ("n" in field.switches) return bookmark.number ?? null;
    return bookmark.text != null ? bookmark.text : null;
  },
  PAGEREF: (field, ctx) => {
    const bookmark = field.args[0] ? ctx.bookmarks?.get(field.args[0]) : undefined;
    if (!bookmark) return null;
    if ("p" in field.switches) return aboveBelow(ctx, bookmark);
    if (bookmark.page == null) return null;
    // The bookmark's own section numbering wins — a cross-section reference
    // formats with the target page's numFmt, not the field's.
    return formatNumber(bookmark.pageFormat ?? ctx.frame?.pageFormat, bookmark.page);
  },
  // NOTEREF — Word's footnote/endnote cross-reference: the reference mark
  // (the note's displayed number, which the host's bookmark scan puts in
  // `text`) instead of the note body. `\p` renders above/below like REF.
  NOTEREF: (field, ctx) => {
    const bookmark = field.args[0] ? ctx.bookmarks?.get(field.args[0]) : undefined;
    if (!bookmark) return null;
    if ("p" in field.switches) return aboveBelow(ctx, bookmark);
    return bookmark.text != null ? bookmark.text : null;
  },
  SEQ: (field, ctx) => {
    const label = field.args[0];
    const ordinal = label ? ctx.sequences?.get(label) : undefined;
    if (ordinal == null) return null;
    const value = formatSeqNumber(field.switches["*"], ordinal);
    // `\s <level>`: prefix the chapter number of the nearest preceding
    // heading at that level (Word's "Include chapter number"). No heading yet
    // → the plain sequence number, Word's no-chapter shape.
    //
    // Word interop deviation: Word renders a chapter-numbered caption as
    // `STYLEREF <level> \s` + separator + `SEQ <label> \s <level>` — the
    // STYLEREF prints the chapter, the SEQ only numbers (and resets). docen
    // caches the chapter prefix inside the SEQ result, so an in-Word field
    // update drops the prefix, and docen's Update All cannot refresh a real
    // Word STYLEREF.
    const level = seqChapterLevel(field.switches.s);
    const chapter = level != null && label ? ctx.chapters?.get(level) : undefined;
    if (chapter == null || !label) return value;
    // Word's default separator is a hyphen when settings carry none.
    return `${chapter}${ctx.captionSeparators?.get(label) ?? "-"}${value}`;
  },
  DOCPROPERTY: (field, ctx) => {
    const name = field.args[0];
    if (!name || !ctx.customProperties) return null;
    const exact = ctx.customProperties.get(name);
    if (exact != null) return exact;
    const lower = name.toLowerCase();
    for (const [key, value] of ctx.customProperties) {
      if (key.toLowerCase() === lower) return value;
    }
    return null;
  },
  INFO: (field, ctx) => {
    const alias = INFO_ALIASES[(field.args[0] ?? "").toLowerCase()];
    const evaluator = alias ? FIELD_EVALUATORS[alias] : undefined;
    if (!evaluator) return null;
    return evaluator({ ...field, name: alias! }, ctx);
  },
};

/** Re-derive a field's value from the document state, or null when this
 *  context cannot provide it (no pagination for PAGE/NUMPAGES, no bookmark
 *  for REF/PAGEREF, unknown instruction) — the caller keeps the cached
 *  result in that case. */
export function resolveField(instruction: string, ctx: FieldContext): string | null {
  const field = parseFieldInstruction(instruction);
  const evaluator = FIELD_EVALUATORS[field.name];
  return evaluator ? evaluator(field, ctx) : null;
}

/** Re-derive the cached value of an updatable field, or null when this field
 *  is not understood in this context — the caller keeps the existing cache. */
export function evaluateField(instruction: string, ctx: FieldContext): string | null {
  return resolveField(instruction, ctx);
}

/** The default instruction of a field name (the dialog's list pick → the
 *  field-code box). */
export const defaultInstruction = (name: string): string =>
  FIELD_CATEGORIES.flatMap((c) => c.fields).find((f) => f.name === name)?.instruction ?? name;

/** The field name of an instruction — the listbox's selection when editing an
 *  existing field ("DATE \@ …" → DATE; unknown → the name still leads). */
export const instructionName = (instruction: string): string =>
  instruction
    .trim()
    .split(/[\s\\]/)[0]
    ?.toUpperCase() ?? "";
