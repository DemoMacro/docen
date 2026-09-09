/**
 * Field catalog + evaluation — the insert/update face of Word's field dialog
 * (插入 → 域). Fields ride the `inlinePassthrough` atom (`attrs.data` carrying
 * the office-open `simpleField`/`complexField`/`formField` branch verbatim), so
 * everything here is plain JSON in, plain JSON out; nothing touches the PM doc.
 *
 * The catalog lists the fields this editor can stand behind: the painter
 * already renders their cached values (layout/project/runs.ts pushField) and
 * `evaluateField` can re-derive those values for 更新域 (F9). Anything the
 * engine neither evaluates nor renders dynamically is left out — a field that
 * shows a stale number is worse than no field.
 */

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
      plain("NUMCHARS"),
      plain("NUMWORDS"),
    ],
  },
  {
    key: "numbering",
    fields: [
      // Dynamic in the painter (resolved per page) — the cached value is just
      // the measuring placeholder Word also caches.
      plain("PAGE"),
      plain("NUMPAGES"),
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

/** What 更新域 needs to re-derive a value. */
export interface FieldContext {
  /** "Now" for DATE/TIME (injected — evaluation must stay testable). */
  now?: Date;
  /** The doc's core properties (creator/title/subject/keywords/comments/
   *  created/modified) backing the document-information fields. */
  core?: Record<string, unknown>;
  /** The document's word and character totals (Word's NUMWORDS/NUMCHARS). */
  words?: number;
  chars?: number;
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

/** Re-derive the cached value of an updatable field, or null when this field
 *  is either dynamic in the painter (PAGE/NUMPAGES) or not understood — the
 *  caller keeps the existing cache in that case. */
export function evaluateField(instruction: string, ctx: FieldContext): string | null {
  const code = instruction.trim();
  const picture = /\\@\s*"([^"]*)"/.exec(code)?.[1];
  const name = (picture ? code.slice(0, code.indexOf(`\\@`)) : code).trim().toUpperCase();
  const now = ctx.now ?? new Date();
  const str = (v: unknown): string | null =>
    typeof v === "string" ? v : typeof v === "number" ? String(v) : null;
  switch (name) {
    case "DATE":
      return formatDate(now, picture ?? "yyyy/M/d");
    case "TIME":
      return formatDate(now, picture ?? "H:mm:ss");
    case "CREATEDATE":
    case "SAVEDATE": {
      const raw = ctx.core?.[name === "CREATEDATE" ? "created" : "modified"];
      const d = typeof raw === "string" || raw instanceof Date ? new Date(raw as string) : null;
      return d && !Number.isNaN(d.getTime()) ? formatDate(d, picture ?? "yyyy/M/d H:mm") : null;
    }
    case "AUTHOR":
      return str(ctx.core?.creator);
    case "TITLE":
      return str(ctx.core?.title);
    case "SUBJECT":
      return str(ctx.core?.subject);
    case "KEYWORDS":
      return str(ctx.core?.keywords);
    case "COMMENTS":
      return str(ctx.core?.description);
    case "NUMWORDS":
      return ctx.words != null ? String(ctx.words) : null;
    case "NUMCHARS":
      return ctx.chars != null ? String(ctx.chars) : null;
    default:
      return null;
  }
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
