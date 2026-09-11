/**
 * Markdown input rules, applied only on the typed leg of the bridge's
 * beforeinput handling (the same path autocorrect rides). The viewless
 * editor has no EditorView, so tiptap's input-rule plugins never fire —
 * these pure functions re-implement the common block/inline conversions
 * and the bridge applies each as one transaction (one undo step).
 *
 * Every rule mirrors CommonMark semantics: block triggers only when the
 * text from the paragraph start IS the trigger sequence; inline fences
 * close only behind a non-empty, non-blank-delimited inner run whose
 * opening delimiter stands at the paragraph start or after whitespace.
 * Like prosemirror-inputrules, inline matching runs against
 * `textBefore + typed` — the bridge hands over the caret text WITHOUT the
 * typed character.
 */

/** One block conversion's target — the bridge maps these onto paragraph
 *  attrs (heading level → HEADING_COMPILE_MAP, quote → IntenseQuote,
 *  bullet/ordered → the flat list attrs). */
export type MarkdownBlockRule =
  | { kind: "heading"; level: number }
  | { kind: "quote" }
  | { kind: "bullet" }
  | { kind: "ordered"; start: number };

/**
 * The block conversion for one typed character, or null. Only a typed
 * space triggers: `textBefore` is the paragraph text before the caret and
 * must BE the whole sequence (`#`..`######`, `>`, `-`/`*`/`+`, `N.`/`N)`).
 */
export function blockRuleOf(typed: string, textBefore: string): MarkdownBlockRule | null {
  if (typed !== " ") return null;
  const heading = /^(#{1,6})$/.exec(textBefore);
  if (heading) return { kind: "heading", level: heading[1]!.length };
  if (textBefore === ">") return { kind: "quote" };
  if (/^[-*+]$/.test(textBefore)) return { kind: "bullet" };
  const ordered = /^(\d+)[.)]$/.exec(textBefore);
  if (ordered && Number(ordered[1]) >= 1) return { kind: "ordered", start: Number(ordered[1]) };
  return null;
}

/**
 * The conversion performed by Enter on a paragraph whose whole text is a
 * markdown fence: an opening code fence (``` or ```lang) turns the
 * paragraph into a Code-styled paragraph carrying the language, a
 * horizontal-rule line (three or more dashes, asterisks or underscores)
 * turns it into a thematic break.
 */
export function enterRuleOf(
  text: string,
): { kind: "code"; language: string | null } | { kind: "hr" } | null {
  const fence = /^```(\S*)$/.exec(text);
  if (fence) return { kind: "code", language: fence[1] || null };
  if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(text)) return { kind: "hr" };
  return null;
}

/** One inline conversion's geometry: the delimiter length (opening and
 *  closing are the same string), the opening delimiter's start and the
 *  inner run's length — both offsets into `textBefore + typed`, the full
 *  text the rule matched. */
export interface MarkdownInlineRule {
  mark: "bold" | "italic" | "strike" | "code";
  openStart: number;
  openLen: number;
  innerLen: number;
}

/** Opening delimiter + inner shape shared by every inline rule: the
 *  delimiter must start the paragraph or follow whitespace (so `a*b*`
 *  stays literal), and the inner run must be non-empty and not
 *  whitespace-bounded (`* a*` / `*a *` don't close). */
function inlineMatch(
  full: string,
  delimiter: string,
): { openStart: number; innerLen: number } | null {
  const escaped = delimiter.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(
    `(^|\\s)${escaped}([^${escaped}\\s](?:[^${escaped}]*[^${escaped}\\s])?)${escaped}$`,
  );
  const m = re.exec(full);
  if (!m) return null;
  return { openStart: m.index + m[1]!.length, innerLen: m[2]!.length };
}

/**
 * The inline conversion completed by one typed character, or null. The
 * typed character IS the last half of the closing delimiter (the second
 * `*` after `**a*`), so the double-delimiter forms are checked before the
 * single ones. `~` and `` ` `` have no doubled form (strike is always
 * `~~`, code always single-backtick).
 */
export function inlineRuleOf(typed: string, textBefore: string): MarkdownInlineRule | null {
  const full = textBefore + typed;
  switch (typed) {
    case "*":
    case "_": {
      const bold = inlineMatch(full, typed.repeat(2));
      if (bold) return { mark: "bold", openLen: 2, ...bold };
      const italic = inlineMatch(full, typed);
      if (italic) return { mark: "italic", openLen: 1, ...italic };
      return null;
    }
    case "~": {
      const strike = inlineMatch(full, "~~");
      if (strike) return { mark: "strike", openLen: 2, ...strike };
      return null;
    }
    case "`": {
      const code = inlineMatch(full, "`");
      if (code) return { mark: "code", openLen: 1, ...code };
      return null;
    }
    default:
      return null;
  }
}

/**
 * True when a typed hyphen must NOT be rewritten by autocorrect: with the
 * input mode on, a hyphen run from the paragraph start is a markdown
 * horizontal-rule in progress (`---` + Enter), so the em-dash
 * substitution must not consume its second hyphen. Off-mode callers never
 * consult this.
 */
export function isHyphenRun(textBefore: string): boolean {
  return /^-+$/.test(textBefore);
}
