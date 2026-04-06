import type {
  JSONContent,
  DeduplicateOptions,
  MatchKind,
  Coverage,
  ParagraphComparison,
  DocumentResult,
  DuplicateMatch,
  SentenceInfo,
  ParagraphInfo,
  SentenceSplitter,
} from "./types";

import { simhash, hammingDistance, levenshteinNormalized, findBestMatch } from "@nlptools/distance";

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------

const DEFAULT_HAMMING_THRESHOLD = 10;
const DEFAULT_THRESHOLD = 0.6;
const DEFAULT_MIN_SENTENCE_LENGTH = 15;

// ---------------------------------------------------------------------------
// Text extraction
// ---------------------------------------------------------------------------

/** Extracts plain text from a Tiptap/ProseMirror JSON node. */
export function extractTextFromNode(node: JSONContent): string {
  if (!node) return "";
  if (node.type === "text" && node.text) return node.text;
  if (node.content && Array.isArray(node.content)) {
    return node.content.map(extractTextFromNode).join("");
  }
  return "";
}

/**
 * Extracts all paragraph/heading text from a Tiptap JSON document.
 *
 * Consecutive `paragraph` nodes are merged when the first does not end
 * with sentence-ending punctuation (。！？.!?). This handles cases where
 * DOCX parsers split one logical paragraph into multiple nodes.
 */
export function extractParagraphs(doc: JSONContent): string[] {
  const paragraphs: string[] = [];
  const SENTENCE_END = /[。！？.!?]$/;

  function processBlock(children: JSONContent[]): void {
    let pending = "";
    for (const child of children) {
      if (!child) continue;

      if (child.type === "paragraph") {
        const text = extractTextFromNode(child);
        if (text && text.trim().length > 0) {
          if (pending && !SENTENCE_END.test(pending.trim())) {
            pending += text;
          } else {
            if (pending) paragraphs.push(pending);
            pending = text;
          }
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
export function splitSentences(text: string): string[] {
  return text
    .split(/(?<=[。！？；\n.!?;])/g)
    .map((s) => s.trim())
    .filter((s) => s.length >= 2);
}

// ---------------------------------------------------------------------------
// Fingerprinting
// ---------------------------------------------------------------------------

/** Generates a SimHash fingerprint for a sentence. Returns null for short sentences. */
export function fingerprintSentence(
  sentence: string,
  minLen = DEFAULT_MIN_SENTENCE_LENGTH,
): bigint | null {
  if (sentence.length < minLen) return null;
  const n = 3;
  const features: string[] = [];
  for (let i = 0; i <= sentence.length - n; i++) {
    features.push(sentence.slice(i, i + n));
  }
  if (features.length === 0) return null;
  return simhash(features);
}

// ---------------------------------------------------------------------------
// Paragraph info
// ---------------------------------------------------------------------------

/** Builds paragraph info with sentence-level SimHash fingerprints. */
export function buildParagraphInfo(
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
  return { text, index, sentences };
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
export function sentenceCoverage(
  paraA: ParagraphInfo,
  paraB: ParagraphInfo,
  hammingThreshold: number,
  levenshteinThreshold: number,
  minFragmentLength: number,
): Coverage {
  const sA = paraA.sentences;
  const sB = paraB.sentences;
  if (sA.length === 0 || sB.length === 0) return { covA: 0, covB: 0 };

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
  if (totalMatchCount === 0) return { covA: 0, covB: 0 };

  const matchedInB = new Set(matchedB.filter((v) => v >= 0));
  return {
    covA: totalMatchCount / sA.length,
    covB: matchedInB.size / sB.length,
  };
}

// ---------------------------------------------------------------------------
// Match classification
// ---------------------------------------------------------------------------

/** Classifies coverage into match kind. */
export function classifyCoverage(covA: number, covB: number): MatchKind {
  const maxCov = Math.max(covA, covB);
  const minCov = Math.min(covA, covB);
  if (maxCov >= 0.8) return "contained";
  if (minCov >= 0.6) return "similar";
  if (maxCov >= 0.3) return "weakOverlap";
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
): DocumentResult {
  const {
    threshold = DEFAULT_THRESHOLD,
    hammingThreshold = DEFAULT_HAMMING_THRESHOLD,
    levenshteinThreshold = DEFAULT_THRESHOLD,
    minSentenceLength = DEFAULT_MIN_SENTENCE_LENGTH,
    splitter = splitSentences,
  } = options;

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
    let bestCoverage: Coverage = { covA: 0, covB: 0 };

    for (const p2 of paras2) {
      const coverage = sentenceCoverage(
        p1,
        p2,
        hammingThreshold,
        levenshteinThreshold,
        minFragmentLength,
      );
      // Select best match: prioritize higher covA, break ties with covB
      if (
        coverage.covA > bestCoverage.covA ||
        (coverage.covA === bestCoverage.covA &&
          coverage.covA > 0 &&
          coverage.covB > bestCoverage.covB)
      ) {
        bestMatch = p2;
        bestCoverage = coverage;
      }
    }

    const similarity = Math.max(bestCoverage.covA, bestCoverage.covB);
    const matchKind =
      similarity < threshold ? "none" : classifyCoverage(bestCoverage.covA, bestCoverage.covB);

    paragraphComparisons.push({
      fromDoc1: { index: p1.index, text: p1.text },
      fromDoc2:
        matchKind !== "none" && bestMatch ? { index: bestMatch.index, text: bestMatch.text } : null,
      coverage: bestCoverage,
      matchKind,
      similarity,
    });
  }

  const coverage =
    paras1.length > 0
      ? paragraphComparisons.reduce((sum, pc) => sum + pc.coverage.covA, 0) / paras1.length
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
    threshold = DEFAULT_THRESHOLD,
    hammingThreshold = DEFAULT_HAMMING_THRESHOLD,
    levenshteinThreshold = DEFAULT_THRESHOLD,
    minSentenceLength = DEFAULT_MIN_SENTENCE_LENGTH,
    splitter = splitSentences,
  } = options;

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
    const similarities: number[] = [];

    for (let j = i + 1; j < paras.length; j++) {
      if (processed.has(j)) continue;
      const coverage = sentenceCoverage(
        paras[i],
        paras[j],
        hammingThreshold,
        levenshteinThreshold,
        minFragmentLength,
      );
      const maxCov = Math.max(coverage.covA, coverage.covB);
      if (maxCov >= threshold) {
        duplicateIndices.push(j);
        similarities.push(maxCov);
        processed.add(j);
      }
    }

    if (duplicateIndices.length > 0) {
      matches.push({
        index: i,
        text: rawParagraphs[i],
        duplicates: duplicateIndices,
        similarities,
      });
      processed.add(i);
    }
  }

  return matches;
}

// ---------------------------------------------------------------------------
// Re-exports
// ---------------------------------------------------------------------------

export { findBestMatch };
