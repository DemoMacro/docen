import { fnv1a, ngrams } from "@nlptools/distance";

// ---------------------------------------------------------------------------
// Winnowing (Schlemer et al. 2003) — verbatim local-match engine. Internal:
// compareDocuments / findDuplicates call winnowLocalMatches per paragraph pair;
// nothing here is part of the public API.
// ---------------------------------------------------------------------------

/** Defaults — k=10, w=6 ⇒ guarantee threshold t = k + w - 1 = 15. Tuned for
 *  CJK: 10-char k-gram is a stable noise floor, w=6 keeps the fingerprint
 *  density ~1 per 3 chars while guaranteeing any ≥15-char shared substring
 *  surfaces at least one matching fingerprint. */
export const DEFAULT_K = 10;
export const DEFAULT_W = 6;

interface Fingerprint {
  hash: number;
  pos: number;
}

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
function winnowFingerprints(text: string, k = DEFAULT_K, w = DEFAULT_W): Fingerprint[] {
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
  const deque: { hash: number; pos: number }[] = [];
  let lastPos = -1;
  for (let i = 0; i < hashes.length; i++) {
    while (deque.length > 0 && deque[deque.length - 1].hash > hashes[i].hash) {
      deque.pop();
    }
    deque.push(hashes[i]);
    while (deque[0].pos <= i - w) deque.shift();
    if (i >= w - 1) {
      const min = deque[0];
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
 * Finds verbatim local overlaps between two paragraphs (the "find copied
 * fragments inside dissimilar text" case whole-paragraph SimHash dilutes).
 *
 * Pipeline: winnow each text once → match fingerprints by hash (a collision is
 * a k-gram seed known identical in both) → `extendSeed` walks each seed out to
 * its full verbatim extent → dedupe (seeds in one fragment extend to the same
 * span). The Winnowing guarantee ⇒ any shared substring of `k + w − 1` chars
 * yields ≥1 fragment. Returns fragments without paragraph indices; the caller
 * wraps them into public LocalMatch records.
 */
export function winnowLocalMatches(
  textA: string,
  textB: string,
  k = DEFAULT_K,
  w = DEFAULT_W,
  minMatch = k + w - 1,
): Fragment[] {
  const fpA = winnowFingerprints(textA, k, w);
  const fpB = winnowFingerprints(textB, k, w);
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
