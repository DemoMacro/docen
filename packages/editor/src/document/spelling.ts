// The spell checker: word tokenization over the PM doc, dictionary lookup
// against the built-in list (extendable at runtime), and edit-distance
// suggestions. Pure computation — the host owns when to run it and what to
// do with the issues (overlay, pane, status).

import type { Node as PMNode } from "@tiptap/pm/model";

import { englishWords } from "./spelling-dictionary";

/** One misspelling: the PM positions of the word (the overlay and the
 *  replace commands work in positions) plus its surface text. */
export interface SpellingIssue {
  from: number;
  to: number;
  word: string;
}

/** A western word — letters with optional inner apostrophes/hyphens. */
const WORD_RE = /^[A-Za-z][A-Za-z'’-]*$/;

/** Session-global extra words (Add to Dictionary) and skipped words (Ignore
 *  All) — shared by every check run so edits don't lose the user's calls. */
const addedWords = new Set<string>();
const ignoredWords = new Set<string>();
/** Ignore Once: exact positions, so only the flagged occurrence is exempt
 *  while the rest of the word keeps its squiggles (Word's per-hit ignore). */
const ignoredOnce = new Set<string>();

/** A word is correct when it (lowercased, or as written — proper nouns sit
 *  outside the lowercase list) hits the built-in list, a host-added word, or
 *  the session's ignored set. */
function known(word: string): boolean {
  const lower = word.toLowerCase();
  return (
    ignoredWords.has(lower) ||
    addedWords.has(lower) ||
    englishWords.has(lower) ||
    englishWords.has(word)
  );
}

/** Spell check one document: every text node is segmented into words and
 *  looked up. Numbers, URLs, CJK runs and anything non-word-like are not
 *  candidates (Word's checker behaves the same for a first pass). Runs marked
 *  "do not check spelling" (w:noProof via the language dialog) are skipped. */
export function checkSpelling(doc: PMNode): SpellingIssue[] {
  const segmenter = new Intl.Segmenter("en", { granularity: "word" });
  const issues: SpellingIssue[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const style = node.marks.find((m) => m.type.name === "textStyle");
    if (style?.attrs.noProof === true) return;
    for (const { segment, index, isWordLike } of segmenter.segment(node.text)) {
      if (!isWordLike || !WORD_RE.test(segment) || known(segment)) continue;
      const from = pos + index;
      if (ignoredOnce.has(`${from}:${segment.toLowerCase()}`)) continue;
      issues.push({ from, to: from + segment.length, word: segment });
    }
  });
  return issues;
}

/** Add a word to the session dictionary (never persisted). */
export function addSpellWord(word: string): void {
  addedWords.add(word.toLowerCase());
}

/** Skip every occurrence of a word for this session. */
export function ignoreSpellWord(word: string): void {
  ignoredWords.add(word.toLowerCase());
}

/** Skip one occurrence (Ignore Once) — keyed on its position, so other hits
 *  of the same word stay flagged. The exemption lapses when the text moves
 *  (an edit reflows the positions), matching how transient Word's is. */
export function ignoreSpellOnce(issue: SpellingIssue): void {
  ignoredOnce.add(`${issue.from}:${issue.word.toLowerCase()}`);
}

/** Damerau-Levenshtein distance (with transpositions) — small words get a
 *  tight radius, longer ones two edits, matching Word's suggestion feel. */
function editDistance(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  let prev2: number[] = [];
  let prev1: number[] = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i, ...Array<number>(n).fill(0)];
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev1[j] + 1, cur[j - 1] + 1, prev1[j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        cur[j] = Math.min(cur[j], prev2[j - 2] + 1);
      }
    }
    prev2 = prev1;
    prev1 = cur;
  }
  return prev1[n];
}

/** Replacement candidates for a misspelling: dictionary words within edit
 *  range, closest first (ties keep the list's frequency order). */
export function spellSuggestions(word: string, limit = 5): string[] {
  const w = word.toLowerCase();
  const radius = w.length <= 4 ? 1 : 2;
  const scored: Array<[string, number]> = [];
  for (const added of addedWords) {
    if (Math.abs(added.length - w.length) <= radius && editDistance(w, added) <= radius) {
      scored.push([added, 0]); // the user's own words win
    }
  }
  for (const cand of englishWords) {
    if (Math.abs(cand.length - w.length) > radius) continue;
    const d = editDistance(w, cand);
    if (d <= radius) scored.push([cand, d]);
  }
  scored.sort((a, b) => a[1] - b[1]);
  return scored.slice(0, limit).map(([w]) => w);
}
