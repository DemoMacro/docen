/**
 * As-you-type autocorrect, applied only where typed characters enter the
 * document (the bridge's beforeinput insertText leg). IME commits, spell-check
 * corrections and pastes ride other insert paths and stay untouched — CJK
 * input therefore never sees any of this by construction.
 *
 * The scope is Word's always-on AutoCorrect behaviors: smart quotes, the
 * built-in word corrections, and the two symbol replacements (dashes,
 * ellipsis) — everything the options dialog will eventually toggle.
 */

/** Straight quote → [opening, closing] curly pair, keyed by the typed char. */
const SMART_QUOTES: Record<string, [string, string]> = {
  '"': ["“", "”"],
  "'": ["‘", "’"],
};

/** Context that makes a typed quote an opening one: nothing behind it,
 *  whitespace, or an opening bracket/quote/dash. Anything else — a word
 *  character, comma, closing quote — makes it closing (Word: don't → don't
 *  with a closing apostrophe, since a letter precedes). */
const OPENS_QUOTE = /^$|[\s([{<—–“‘'-]$/u;

/** The typed quote's curly replacement, or null for anything else. */
export function smartQuoteOf(typed: string, charBefore: string): string | null {
  const pair = SMART_QUOTES[typed];
  if (!pair) return null;
  return OPENS_QUOTE.test(charBefore) ? pair[0]! : pair[1]!;
}

/** Word's built-in correction table (common subset). */
const CORRECTIONS: Record<string, string> = {
  teh: "the",
  adn: "and",
  taht: "that",
  thier: "their",
  recieve: "receive",
  seperate: "separate",
  occured: "occurred",
  wich: "which",
  writen: "written",
  beleive: "believe",
  freind: "friend",
  definately: "definitely",
};

/** Boundary characters that trigger a word correction behind them. */
const WORD_BOUNDARY = /[\s,.!?;:'")\]}]/u;

/** Preserve the corrected word's case shape: lower → lower, ALL CAPS → ALL
 *  CAPS, Capitalized → Capitalized (Word's initial-caps rule). */
function matchCase(replacement: string, source: string): string {
  if (source === source.toUpperCase() && source !== source.toLowerCase()) {
    return replacement.toUpperCase();
  }
  if (/^\p{Lu}/u.test(source)) {
    return replacement.charAt(0).toUpperCase() + replacement.slice(1);
  }
  return replacement;
}

/** The correction for `word`, or null. */
export function correctionOf(word: string): string | null {
  const fixed = CORRECTIONS[word.toLowerCase()];
  return fixed ? matchCase(fixed, word) : null;
}

/**
 * The autocorrect result of one typed character.
 *
 * Returns the replacement text (the typed char included — corrections ride
 * their boundary character) plus `back`, how many characters before the
 * caret it overwrites. The caller applies
 * `tr.insertText(text, from - back, to)`. Null when the character needs no
 * rewrite.
 *
 * `textBefore` is the paragraph text before the caret — quotes read one
 * character back, word corrections read to the last boundary.
 */
export function autocorrectOf(
  typed: string,
  textBefore: string,
): {
  text: string;
  back: number;
} | null {
  const quote = smartQuoteOf(typed, textBefore.slice(-1));
  if (quote) return { text: quote, back: 0 };

  // Dashes/ellipsis: the typed char completes a run in the text behind. The
  // second hyphen of a pair becomes Word's em dash (the first is already
  // behind the caret); the third dot completes "...".
  if (typed === "-" && textBefore.endsWith("-")) return { text: "—", back: 1 };
  if (typed === "." && textBefore.endsWith("..")) return { text: "…", back: 2 };

  // Word correction: the typed char is a boundary, the word behind it may
  // need fixing. The boundary trails the corrected word — "teh " becomes
  // "the " with the typed space kept as the word separator.
  if (!WORD_BOUNDARY.test(typed) && typed !== "\n") return null;
  const word = /[\p{L}\p{N}]+$/u.exec(textBefore)?.[0];
  if (!word) return null;
  const fixed = correctionOf(word);
  if (!fixed) return null;
  return { text: fixed + typed, back: word.length };
}
