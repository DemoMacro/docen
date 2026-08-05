import { fnv1a, ngrams } from "@nlptools/distance";

import type { Fingerprint } from "./types";

// ---------------------------------------------------------------------------
// Winnowing (Schlemer et al. 2003) — verbatim local-match engine. Internal:
// compareDocuments / findDuplicates precompute fingerprints once per paragraph
// (ParagraphInfo.winnowFp) and pair them via winnowLocalMatchesFromFingerprints;
// nothing here is part of the public API.
// ---------------------------------------------------------------------------

/** Defaults — k=10, w=4 ⇒ guarantee threshold t = k + w - 1 = 13.
 *
 *  The 13-char guarantee aligns with the Chinese academic plagiarism-check
 *  industry standard: CNKI (知网), Wanfang (万方), VIP (维普), and PaperPass all
 *  flag "13 consecutive matching characters" as a duplicate. docen's Winnowing
 *  layer catches verbatim copies *inside* an otherwise-dissimilar paragraph
 *  that whole-paragraph SimHash dilutes, so the guarantee hugs that industry
 *  floor — any shared substring of 13+ chars within a paragraph pair is
 *  reported. The 10-char k-gram stays below it as the noise floor (sub-10-char
 *  runs don't even seed); minMatchLength defaults to t=13, so only 13+ char
 *  spans are reported, matching the industry rule precisely. (Schleimer et
 *  al.'s k≈50 for whole-document English prose optimizes for precision over a
 *  large corpus; docen's per-paragraph CJK role calls for the shorter,
 *  industry-aligned threshold.) */
export const DEFAULT_K = 10;
export const DEFAULT_W = 4;

/** A verbatim overlap between two texts (paragraph-pair level; paraIndex is
 *  filled by the caller when wrapping into the public LocalMatch). */
export interface Fragment {
  startA: number;
  endA: number;
  startB: number;
  endB: number;
  length: number;
  text: string;
}

/**
 * Winnowing fingerprints of `text`: hash each overlapping k-gram, then keep
 * the minimum hash in each window of w (rightmost on ties; robust variant
 * keeps the prior selection while it stays in-window). Guarantees any shared
 * substring of length ≥ k + w − 1 yields at least one matching fingerprint in
 * both texts.
 *
 * Built on @nlptools/distance's `ngrams` + `fnv1a`; only the windowed-min
 * selection is docen's own — the library ships the parts but not the winnow
 * step itself.
 */
export function winnowFingerprints(text: string, k = DEFAULT_K, w = DEFAULT_W): Fingerprint[] {
  if (k < 1) throw new Error(`winnowFingerprints: k must be >= 1, got ${k}`);
  if (w < 1) throw new Error(`winnowFingerprints: w must be >= 1, got ${w}`);
  const grams = ngrams(text, k);
  if (grams.length === 0) return [];
  const hashes: { hash: number; pos: number }[] = grams.map((g, i) => ({
    hash: fnv1a(g),
    pos: i,
  }));
  return winnowSelect(hashes, w);
}

/**
 * Monotonic-deque winnow: select the minimum hash in each window of w, keeping
 * the prior selection while it stays in-window and minimal (robust variant —
 * collapses repeated selections on low-entropy runs). Returns fingerprints in
 * source order.
 */
function winnowSelect(hashes: { hash: number; pos: number }[], w: number): Fingerprint[] {
  const fingerprints: Fingerprint[] = [];
  // Monotonic deque of sliding-window-minimum candidates, stored as object
  // references with a head index. push/pop act on the tail (O(1)); advancing
  // the head index evicts off-window entries (O(1)) — no Array.shift(), which
  // is O(queue length) and would make the loop O(n·w). Entries before `head`
  // stay in the array (lazy) but are never read; the array is per-call and
  // short-lived, so the slack is harmless.
  const deque: { hash: number; pos: number }[] = [];
  let head = 0;
  let lastPos = -1;
  for (let i = 0; i < hashes.length; i++) {
    while (deque.length > head && deque[deque.length - 1].hash > hashes[i].hash) {
      deque.pop();
    }
    deque.push(hashes[i]);
    while (deque[head].pos <= i - w) head++;
    if (i >= w - 1) {
      const min = deque[head];
      if (min.pos !== lastPos) {
        fingerprints.push({ hash: min.hash, pos: min.pos });
        lastPos = min.pos;
      }
    }
  }
  return fingerprints;
}

/**
 * Grows a seed (a matched k-gram at `posA`/`posB`) outward to its full verbatim
 * extent by walking both texts char-by-char until they diverge. Turns a k-char
 * anchor into a precisely bounded fragment of any length.
 */
function extendSeed(
  textA: string,
  textB: string,
  posA: number,
  posB: number,
  k: number,
): { startA: number; endA: number; startB: number; endB: number } {
  let startA = posA;
  let startB = posB;
  while (startA > 0 && startB > 0 && textA[startA - 1] === textB[startB - 1]) {
    startA--;
    startB--;
  }
  let endA = posA + k;
  let endB = posB + k;
  while (endA < textA.length && endB < textB.length && textA[endA] === textB[endB]) {
    endA++;
    endB++;
  }
  return { startA, endA, startB, endB };
}

/**
 * Finds verbatim local overlaps between two paragraphs from precomputed
 * fingerprints (the "find copied fragments inside dissimilar text" case
 * whole-paragraph SimHash dilutes). Callers pairing many paragraphs should
 * precompute fingerprints once per paragraph ({@link winnowFingerprints}) and
 * reuse them here — fingerprinting every pair is O(P²) winnows, this turns it
 * into O(P) winnows + O(P²) cheap hash lookups.
 *
 * Pipeline: match fingerprints by hash (a collision is a k-gram seed known
 * identical in both) → `extendSeed` walks each seed out to its full verbatim
 * extent → dedupe (seeds in one fragment extend to the same span). The
 * Winnowing guarantee ⇒ any shared substring of `k + w − 1` chars yields ≥1
 * fragment. Returns fragments without paragraph indices; the caller wraps them
 * into public LocalMatch records.
 */
export function winnowLocalMatchesFromFingerprints(
  fpA: Fingerprint[],
  fpB: Fingerprint[],
  textA: string,
  textB: string,
  k: number,
  minMatch: number,
): Fragment[] {
  if (fpA.length === 0 || fpB.length === 0) return [];

  const aIndex = new Map<number, number[]>();
  for (const fp of fpA) {
    const list = aIndex.get(fp.hash);
    if (list) list.push(fp.pos);
    else aIndex.set(fp.hash, [fp.pos]);
  }

  const fragments: Fragment[] = [];
  // Dedup key: (startA, startB). Seeds inside one fragment extend to identical
  // spans, so this collapses them to a single report.
  const seen = new Set<string>();

  for (const fb of fpB) {
    const positionsA = aIndex.get(fb.hash);
    if (!positionsA) continue;
    for (const posA of positionsA) {
      const { startA, endA, startB, endB } = extendSeed(textA, textB, posA, fb.pos, k);
      const length = endA - startA;
      if (length < minMatch) continue;
      const key = `${startA}:${startB}`;
      if (seen.has(key)) continue;
      seen.add(key);
      fragments.push({ startA, endA, startB, endB, length, text: textA.slice(startA, endA) });
    }
  }

  return fragments;
}

/** Convenience wrapper: fingerprints both texts on the fly, then delegates to
 *  {@link winnowLocalMatchesFromFingerprints}. Use for one-off pairs; for
 *  repeated pairing (one paragraph vs. many), precompute fingerprints with
 *  {@link winnowFingerprints} and call the From-Fingerprints form to avoid
 *  recomputation. */
export function winnowLocalMatches(
  textA: string,
  textB: string,
  k = DEFAULT_K,
  w = DEFAULT_W,
  minMatch = k + w - 1,
): Fragment[] {
  const fpA = winnowFingerprints(textA, k, w);
  const fpB = winnowFingerprints(textB, k, w);
  return winnowLocalMatchesFromFingerprints(fpA, fpB, textA, textB, k, minMatch);
}
