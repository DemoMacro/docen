import { hammingDistance, levenshteinNormalized, simhash } from "@nlptools/distance";

import type {
  Coverage,
  DeduplicateOptions,
  DocumentComparison,
  DuplicateMatch,
  JSONContent,
  LocalMatch,
  LocalMatchConfig,
  MatchKind,
  ParagraphComparison,
  ParagraphInfo,
  SentenceInfo,
  SentenceSplitter,
} from "./types";
import { DEFAULT_K, DEFAULT_W, winnowLocalMatches, type Fragment } from "./winnowing";

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------

const DEFAULT_HAMMING_THRESHOLD = 10;
const DEFAULT_SIMILARITY_THRESHOLD = 0.6;
const DEFAULT_MIN_SENTENCE_LENGTH = 15;

// ---------------------------------------------------------------------------
// Text extraction
// ---------------------------------------------------------------------------

/** Extracts plain text from a Tiptap/ProseMirror JSON node. */
function extractTextFromNode(node: JSONContent): string {
  if (!node) return "";
  if (node.type === "text" && node.text) return node.text;
  if (node.content && Array.isArray(node.content)) {
    return node.content.map(extractTextFromNode).join("");
  }
  return "";
}

// ---------------------------------------------------------------------------
// Drawing nodes — DOCX text boxes (wpsShape) and drawing groups (wpgGroup) are
// inline nodes whose text body must be pulled out as standalone paragraphs
// rather than merged into the host paragraph. Hard-coded against @docen/docx
// node shapes; this package does not depend on @docen/docx.
// ---------------------------------------------------------------------------

/** wpsShape: standalone text box (inline, content is block+ — paragraphs).
 *  wpgGroup: drawing group (atom — full model lives in attrs.wpgGroup). */
type DrawingNodeType = "wpsShape" | "wpgGroup";

function isDrawingNode(
  node: JSONContent | undefined | null,
): node is JSONContent & { type: DrawingNodeType } {
  return node?.type === "wpsShape" || node?.type === "wpgGroup";
}

/** Text of a node's inline children, skipping drawing nodes — their text is
 *  collected separately so a text box's content stays a standalone paragraph
 *  and never bleeds into its host. */
function directText(node: JSONContent | undefined | null): string {
  if (!node) return "";
  if (node.type === "text" && node.text) return node.text;
  if (isDrawingNode(node)) return "";
  if (Array.isArray(node.content)) return node.content.map(directText).join("");
  return "";
}

/** office-open ParagraphOptions → plain text. A run is `{ text: "..." }` or a
 *  bare string; non-text run fields (style, break) contribute nothing. */
function paragraphOptionsText(para: unknown): string {
  if (typeof para === "string") return para;
  if (!para || typeof para !== "object") return "";
  const children = (para as { children?: unknown }).children;
  if (!Array.isArray(children)) return "";
  return children
    .map((run) => {
      if (typeof run === "string") return run;
      const text = (run as { text?: unknown } | null)?.text;
      return typeof text === "string" ? text : "";
    })
    .join("");
}

/** Extracts text-box paragraphs from a wpg group's children JSON
 *  (`attrs.wpgGroup.children`). Each child is a wps shape
 *  (`{ type: "wps", data: { children: ParagraphOptions[] } }`) or a nested wpg
 *  (`{ type: "wpg", children: [...] }`); pic children carry no text. */
function extractWpgChildrenText(children: unknown, out: string[]): void {
  if (!Array.isArray(children)) return;
  for (const child of children) {
    if (!child || typeof child !== "object") continue;
    const c = child as { type?: string; data?: { children?: unknown[] }; children?: unknown[] };
    if (c.type === "wpg") {
      extractWpgChildrenText(c.children, out);
    } else if (c.type === "wps" && Array.isArray(c.data?.children)) {
      for (const para of c.data!.children) {
        const text = paragraphOptionsText(para);
        if (text.trim()) out.push(text);
      }
    }
  }
}

/** Walks a node's descendants collecting text-box paragraphs from any
 *  wpsShape (text body is PM `content`) or wpgGroup (text body is attrs JSON).
 *  Each located paragraph is pushed as a standalone entry. */
function collectTextBoxText(node: JSONContent, out: string[]): void {
  if (!node) return;
  if (node.type === "wpsShape") {
    if (Array.isArray(node.content)) {
      for (const para of node.content) {
        const text = directText(para);
        if (text.trim()) out.push(text);
      }
    }
    return;
  }
  if (node.type === "wpgGroup") {
    const groupChildren = (node.attrs?.wpgGroup as { children?: unknown } | undefined)?.children;
    extractWpgChildrenText(groupChildren, out);
    return;
  }
  if (Array.isArray(node.content)) {
    for (const child of node.content) collectTextBoxText(child, out);
  }
}

/**
 * Extracts all paragraph/heading text from a Tiptap JSON document.
 *
 * Consecutive `paragraph` nodes are merged when the first does not end
 * with sentence-ending punctuation (。！？.!?). This handles cases where
 * DOCX parsers split one logical paragraph into multiple nodes.
 *
 * DOCX text-box content (wpsShape/wpgGroup) is extracted as standalone
 * paragraphs — a text box's body never merges into its host paragraph.
 */
export function extractParagraphs(doc: JSONContent): string[] {
  const paragraphs: string[] = [];
  const SENTENCE_END = /[。！？.!?]$/;

  function processBlock(children: JSONContent[]): void {
    let pending = "";
    for (const child of children) {
      if (!child) continue;

      if (child.type === "paragraph") {
        const text = directText(child);
        const textBoxes: string[] = [];
        collectTextBoxText(child, textBoxes);
        if (text && text.trim().length > 0) {
          if (pending && !SENTENCE_END.test(pending.trim())) {
            pending += text;
          } else {
            if (pending) paragraphs.push(pending);
            pending = text;
          }
        }
        // Text-box content stands alone — flush pending first so it never
        // merges with the host paragraph's accumulated text.
        for (const tb of textBoxes) {
          if (!tb.trim()) continue;
          if (pending) paragraphs.push(pending);
          pending = "";
          paragraphs.push(tb);
        }
      } else if (child.type === "heading") {
        if (pending) paragraphs.push(pending);
        pending = "";
        const text = extractTextFromNode(child);
        if (text && text.trim().length > 0) paragraphs.push(text);
      } else {
        if (pending) paragraphs.push(pending);
        pending = "";
        if (child.content && Array.isArray(child.content)) {
          processBlock(child.content);
        }
      }
    }
    if (pending) paragraphs.push(pending);
  }

  if (doc.content && Array.isArray(doc.content)) {
    processBlock(doc.content);
  }
  return paragraphs;
}

/** Splits text into sentences. Chinese-aware. */
function splitSentences(text: string): string[] {
  return text
    .split(/(?<=[。！？；\n.!?;])/g)
    .map((s) => s.trim())
    .filter((s) => s.length >= 2);
}

// ---------------------------------------------------------------------------
// Fingerprinting
// ---------------------------------------------------------------------------

/** Builds a SimHash fingerprint from character trigrams, skipping text shorter
 *  than `minLen`. Shared by sentence- and paragraph-level fingerprinting so the
 *  feature extraction stays in one place. Uses @nlptools/distance's simhash. */
function simhashTrigrams(text: string, minLen: number): bigint | null {
  if (text.length < minLen) return null;
  const n = 3;
  const features: string[] = [];
  for (let i = 0; i <= text.length - n; i++) {
    features.push(text.slice(i, i + n));
  }
  if (features.length === 0) return null;
  return simhash(features);
}

/** Generates a SimHash fingerprint for a sentence. Returns null for short sentences. */
function fingerprintSentence(
  sentence: string,
  minLen = DEFAULT_MIN_SENTENCE_LENGTH,
): bigint | null {
  return simhashTrigrams(sentence, minLen);
}

/** Generates a SimHash fingerprint for a whole paragraph (trigram-based), used
 *  for O(1) paragraph-pair prescreening before the expensive sentence-level
 *  matching. Returns null for short paragraphs. */
function fingerprintParagraph(text: string, minLen = DEFAULT_MIN_SENTENCE_LENGTH): bigint | null {
  return simhashTrigrams(text, minLen);
}

// ---------------------------------------------------------------------------
// Paragraph info
// ---------------------------------------------------------------------------

/** Builds paragraph info with sentence-level SimHash fingerprints. */
function buildParagraphInfo(
  text: string,
  index: number,
  splitter: SentenceSplitter = splitSentences,
  minSentenceLength = DEFAULT_MIN_SENTENCE_LENGTH,
): ParagraphInfo {
  const raw = splitter(text);
  const sentences: SentenceInfo[] = raw.map((s) => ({
    text: s,
    fingerprint: fingerprintSentence(s, minSentenceLength),
  }));
  return {
    text,
    index,
    sentences,
    fingerprint: fingerprintParagraph(text, minSentenceLength),
  };
}

// ---------------------------------------------------------------------------
// Paragraph-pair prescreening
// ---------------------------------------------------------------------------

/** Cheap O(1) paragraph-pair prescreen using @nlptools/distance's
 *  hammingDistance. Skips the expensive sentence-level matching when both
 *  paragraphs carry fingerprints and their SimHash distance exceeds the
 *  threshold. Short paragraphs (null fingerprint) bypass prescreening and fall
 *  through to sentence-level comparison. */
function isParagraphCandidate(
  fpA: bigint | null,
  fpB: bigint | null,
  hammingThreshold: number,
): boolean {
  if (fpA === null || fpB === null) return true;
  return hammingDistance(fpA, fpB) <= hammingThreshold;
}

// ---------------------------------------------------------------------------
// Local-match (Winnowing) integration
// ---------------------------------------------------------------------------

/** Resolves the localMatch option into concrete parameters, or null when
 *  disabled. Tri-state: omitted/true → defaults, false → off, object → tuned. */
function resolveLocalMatchConfig(
  option: boolean | LocalMatchConfig | undefined,
): { k: number; w: number; minMatch: number } | null {
  if (option === false) return null;
  if (option === undefined || option === true) {
    return { k: DEFAULT_K, w: DEFAULT_W, minMatch: DEFAULT_K + DEFAULT_W - 1 };
  }
  const k = option.kgramLength ?? DEFAULT_K;
  const w = option.windowSize ?? DEFAULT_W;
  const minMatch = option.minMatchLength ?? k + w - 1;
  return { k, w, minMatch };
}

/** Wraps a paragraph-pair fragment (no paragraphIndex) into a public LocalMatch. */
function toLocalMatch(frag: Fragment, paraA: number, paraB: number): LocalMatch {
  return {
    fromDoc1: { paragraphIndex: paraA, start: frag.startA, end: frag.endA, text: frag.text },
    fromDoc2: { paragraphIndex: paraB, start: frag.startB, end: frag.endB, text: frag.text },
    length: frag.length,
  };
}

// ---------------------------------------------------------------------------
// Sentence-level coverage
// ---------------------------------------------------------------------------

/**
 * Computes bidirectional coverage between two paragraphs using
 * sentence-level matching.
 *
 * Two-phase matching:
 * 1. SimHash hamming distance for fast candidate screening (fingerprinted sentences)
 * 2. Levenshtein normalized similarity for precise verification (all sentences)
 *
 * No containment fallback — if no sentence-level match is found, coverage is 0.
 */
function sentenceCoverage(
  paraA: ParagraphInfo,
  paraB: ParagraphInfo,
  hammingThreshold: number,
  levenshteinThreshold: number,
  minFragmentLength: number,
): Coverage {
  const sA = paraA.sentences;
  const sB = paraB.sentences;
  if (sA.length === 0 || sB.length === 0) return { coverageA: 0, coverageB: 0 };

  const matchedB = new Int32Array(sA.length).fill(-1);
  const usedB = new Uint8Array(sB.length);

  // Phase 1: SimHash screening → Levenshtein verification (fingerprinted sentences)
  for (let i = 0; i < sA.length; i++) {
    const fpA = sA[i].fingerprint;
    if (fpA === null) continue;
    for (let j = 0; j < sB.length; j++) {
      if (usedB[j] || sB[j].fingerprint === null) continue;
      if (hammingDistance(fpA, sB[j].fingerprint!) <= hammingThreshold) {
        if (levenshteinNormalized(sA[i].text, sB[j].text) >= levenshteinThreshold) {
          matchedB[i] = j;
          usedB[j] = 1;
          break;
        }
      }
    }
  }

  // Phase 2: Direct Levenshtein for unmatched sentences (including short ones)
  for (let i = 0; i < sA.length; i++) {
    if (matchedB[i] >= 0) continue;
    if (sA[i].text.length < minFragmentLength) continue;
    let bestJ = -1;
    let bestScore = 0;
    for (let j = 0; j < sB.length; j++) {
      if (usedB[j]) continue;
      const score = levenshteinNormalized(sA[i].text, sB[j].text);
      if (score > bestScore) {
        bestScore = score;
        bestJ = j;
      }
    }
    if (bestJ >= 0 && bestScore >= levenshteinThreshold) {
      matchedB[i] = bestJ;
      usedB[bestJ] = 1;
    }
  }

  const totalMatchCount = matchedB.filter((v) => v >= 0).length;
  if (totalMatchCount === 0) return { coverageA: 0, coverageB: 0 };

  const matchedInB = new Set(matchedB.filter((v) => v >= 0));
  return {
    coverageA: totalMatchCount / sA.length,
    coverageB: matchedInB.size / sB.length,
  };
}

// ---------------------------------------------------------------------------
// Match classification
// ---------------------------------------------------------------------------

/** Classifies coverage into match kind. `similarityThreshold` is the noise
 *  floor for partial — defaults to 0.3, but compareDocuments passes the user's
 *  threshold so the "max >= threshold" contract holds even when it is < 0.3. */
function classifyCoverage(
  coverageA: number,
  coverageB: number,
  similarityThreshold = 0.3,
): MatchKind {
  const maxCov = Math.max(coverageA, coverageB);
  const minCov = Math.min(coverageA, coverageB);
  if (maxCov >= 0.8) return "contained";
  if (minCov >= 0.6) return "similar";
  if (maxCov >= similarityThreshold) return "partial";
  return "none";
}

// ---------------------------------------------------------------------------
// Main API
// ---------------------------------------------------------------------------

/**
 * Calculates similarity between two texts using Levenshtein normalized distance.
 * @returns number between 0 (completely different) and 1 (identical)
 */
export function calculateSimilarity(text1: string, text2: string): number {
  const t1 = text1.trim();
  const t2 = text2.trim();
  if (!t1 && !t2) return 1;
  if (!t1 || !t2) return 0;
  return levenshteinNormalized(t1, t2);
}

/**
 * Compares two Tiptap JSON documents.
 * Returns per-paragraph comparisons and an overall document coverage score.
 */
export function compareDocuments(
  doc1: JSONContent,
  doc2: JSONContent,
  options: DeduplicateOptions = {},
): DocumentComparison {
  const {
    similarityThreshold = DEFAULT_SIMILARITY_THRESHOLD,
    hammingThreshold = DEFAULT_HAMMING_THRESHOLD,
    levenshteinThreshold = DEFAULT_SIMILARITY_THRESHOLD,
    minSentenceLength = DEFAULT_MIN_SENTENCE_LENGTH,
    splitter = splitSentences,
  } = options;
  const localMatchCfg = resolveLocalMatchConfig(options.localMatch);

  const rawParas1 = extractParagraphs(doc1);
  const rawParas2 = extractParagraphs(doc2);
  const paras1 = rawParas1.map((text, i) =>
    buildParagraphInfo(text, i, splitter, minSentenceLength),
  );
  const paras2 = rawParas2.map((text, i) =>
    buildParagraphInfo(text, i, splitter, minSentenceLength),
  );
  const minFragmentLength = Math.max(2, Math.floor(minSentenceLength / 4));

  const paragraphComparisons: ParagraphComparison[] = [];

  for (const p1 of paras1) {
    let bestMatch: ParagraphInfo | null = null;
    let bestCoverage: Coverage = { coverageA: 0, coverageB: 0 };
    // Verbatim fragments gathered across ALL candidate p2 — a copied fragment
    // can sit in a paragraph that isn't the sentence-coverage best match.
    const verbatimMatches: LocalMatch[] = [];

    for (const p2 of paras2) {
      // Sentence matching is gated by SimHash prescreen (skips O(s²) for far
      // pairs); Winnowing always runs, since a verbatim copy can sit inside a
      // SimHash-distant pair that sentence matching would skip entirely.
      const coverage = isParagraphCandidate(p1.fingerprint, p2.fingerprint, hammingThreshold)
        ? sentenceCoverage(p1, p2, hammingThreshold, levenshteinThreshold, minFragmentLength)
        : { coverageA: 0, coverageB: 0 };
      if (localMatchCfg) {
        const frags = winnowLocalMatches(
          p1.text,
          p2.text,
          localMatchCfg.k,
          localMatchCfg.w,
          localMatchCfg.minMatch,
        );
        for (const frag of frags) verbatimMatches.push(toLocalMatch(frag, p1.index, p2.index));
      }
      // Select best match: prioritize higher coverageA, break ties with coverageB
      if (
        coverage.coverageA > bestCoverage.coverageA ||
        (coverage.coverageA === bestCoverage.coverageA &&
          coverage.coverageA > 0 &&
          coverage.coverageB > bestCoverage.coverageB)
      ) {
        bestMatch = p2;
        bestCoverage = coverage;
      }
    }

    const similarity = Math.max(bestCoverage.coverageA, bestCoverage.coverageB);
    // Winnowing upgrade: verbatim fragments surface copied substrings even when
    // whole-paragraph similarity is below the noise floor (the "a hundred-char
    // paragraph with a dozen copied characters" case SimHash dilutes).
    const matchKind: MatchKind =
      similarity < similarityThreshold
        ? verbatimMatches.length > 0
          ? "partial"
          : "none"
        : classifyCoverage(bestCoverage.coverageA, bestCoverage.coverageB, similarityThreshold);

    paragraphComparisons.push({
      fromDoc1: { index: p1.index, text: p1.text },
      fromDoc2:
        matchKind !== "none" && bestMatch ? { index: bestMatch.index, text: bestMatch.text } : null,
      coverage: bestCoverage,
      matchKind,
      similarity,
      verbatimMatches,
    });
  }

  const totalSentences = paras1.reduce((sum, p) => sum + p.sentences.length, 0);
  const coverage =
    totalSentences > 0
      ? paragraphComparisons.reduce(
          (sum, pc, i) => sum + pc.coverage.coverageA * paras1[i].sentences.length,
          0,
        ) / totalSentences
      : 0;

  return { paragraphs: paragraphComparisons, coverage };
}

/**
 * Finds duplicate/similar paragraphs within a single document.
 */
export function findDuplicates(
  doc: JSONContent,
  options: DeduplicateOptions = {},
): DuplicateMatch[] {
  const {
    similarityThreshold = DEFAULT_SIMILARITY_THRESHOLD,
    hammingThreshold = DEFAULT_HAMMING_THRESHOLD,
    levenshteinThreshold = DEFAULT_SIMILARITY_THRESHOLD,
    minSentenceLength = DEFAULT_MIN_SENTENCE_LENGTH,
    splitter = splitSentences,
  } = options;
  const localMatchCfg = resolveLocalMatchConfig(options.localMatch);

  const rawParagraphs = extractParagraphs(doc);
  const paras = rawParagraphs.map((text, i) =>
    buildParagraphInfo(text, i, splitter, minSentenceLength),
  );
  const minFragmentLength = Math.max(2, Math.floor(minSentenceLength / 4));

  const matches: DuplicateMatch[] = [];
  const processed = new Set<number>();

  for (let i = 0; i < paras.length; i++) {
    if (processed.has(i)) continue;

    const duplicateIndices: number[] = [];
    const similarityScores: number[] = [];
    const verbatimMatches: LocalMatch[] = [];

    for (let j = i + 1; j < paras.length; j++) {
      if (processed.has(j)) continue;
      // Sentence matching is gated by SimHash prescreen (skips O(s²) for far
      // pairs); Winnowing always runs so verbatim copies inside SimHash-distant
      // pairs are still caught.
      const coverage = isParagraphCandidate(
        paras[i].fingerprint,
        paras[j].fingerprint,
        hammingThreshold,
      )
        ? sentenceCoverage(
            paras[i],
            paras[j],
            hammingThreshold,
            levenshteinThreshold,
            minFragmentLength,
          )
        : { coverageA: 0, coverageB: 0 };
      const frags = localMatchCfg
        ? winnowLocalMatches(
            paras[i].text,
            paras[j].text,
            localMatchCfg.k,
            localMatchCfg.w,
            localMatchCfg.minMatch,
          )
        : [];
      const maxCov = Math.max(coverage.coverageA, coverage.coverageB);
      // A pair counts as duplicate when sentence-level similarity clears the
      // threshold OR a verbatim fragment survives (Winnowing — copied text
      // inside otherwise-different paragraphs).
      if (maxCov >= similarityThreshold || frags.length > 0) {
        duplicateIndices.push(j);
        similarityScores.push(maxCov);
        for (const frag of frags) verbatimMatches.push(toLocalMatch(frag, i, j));
        processed.add(j);
      }
    }

    if (duplicateIndices.length > 0) {
      matches.push({
        index: i,
        text: rawParagraphs[i],
        duplicateIndices,
        similarityScores,
        verbatimMatches,
      });
      processed.add(i);
    }
  }

  return matches;
}
