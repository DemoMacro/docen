import type { JSONContent } from "@tiptap/core";

export type { JSONContent };

// ---------------------------------------------------------------------------
// Match classification
// ---------------------------------------------------------------------------

/** Classification of text similarity relationship. */
export type MatchKind = "contained" | "similar" | "partial" | "none";

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/** Sentence splitter function. */
export type SentenceSplitter = (text: string) => string[];

/** Tuning for verbatim local-match (Winnowing). */
export interface LocalMatchConfig {
  /** k-gram length — noise floor; matches shorter than this are invisible. @default 10 */
  kgramLength?: number;
  /** Winnowing window size; guarantee threshold t = kgramLength + windowSize - 1. @default 6 */
  windowSize?: number;
  /** Minimum reported fragment length. @default kgramLength + windowSize - 1 (the guarantee threshold) */
  minMatchLength?: number;
}

/** Configuration options for deduplication. */
export interface DeduplicateOptions {
  /** Minimum similarity (0-1) for duplicate detection. @default 0.6 */
  similarityThreshold?: number;
  /** SimHash hamming distance threshold for candidate screening. @default 10 */
  hammingThreshold?: number;
  /** Levenshtein similarity threshold for fine-grained verification. @default 0.6 */
  levenshteinThreshold?: number;
  /** Minimum sentence length for SimHash fingerprinting and Phase 2 matching. @default 15 */
  minSentenceLength?: number;
  /** Custom sentence splitter. @default splitSentences (Chinese & English aware) */
  splitter?: SentenceSplitter;
  /**
   * Verbatim local-match (Winnowing) — catches copied fragments inside
   * dissimilar text that whole-paragraph SimHash dilutes (the "a hundred-char
   * paragraph with a dozen copied characters" case). Pass `false` to disable,
   * or a {@link LocalMatchConfig} to tune; default enabled (k=10, w=6 ⇒ 15-char
   * guarantee).
   * @default true
   */
  localMatch?: boolean | LocalMatchConfig;
}

// ---------------------------------------------------------------------------
// Internal types
// ---------------------------------------------------------------------------

/** Sentence with optional SimHash fingerprint. */
export interface SentenceInfo {
  text: string;
  fingerprint: bigint | null;
}

/** Paragraph with sentence-level information. */
export interface ParagraphInfo {
  text: string;
  index: number;
  sentences: SentenceInfo[];
  /** Paragraph-level SimHash for O(1) pair prescreening. Null if too short. */
  fingerprint: bigint | null;
}

// ---------------------------------------------------------------------------
// Result types
// ---------------------------------------------------------------------------

/** Bidirectional coverage between two paragraphs. */
export interface Coverage {
  /** Proportion of document 1's sentences matched in document 2 (0-1) */
  coverageA: number;
  /** Proportion of document 2's sentences matched in document 1 (0-1) */
  coverageB: number;
}

/** Paragraph-level comparison result between two documents. */
export interface ParagraphComparison {
  /** Paragraph from document 1 */
  fromDoc1: { index: number; text: string };
  /** Best matching paragraph from document 2 (null if no match) */
  fromDoc2: { index: number; text: string } | null;
  /** Bidirectional sentence-level coverage */
  coverage: Coverage;
  /** Classification of the relationship */
  matchKind: MatchKind;
  /** Overall similarity score (max of coverageA, coverageB) */
  similarity: number;
  /** Verbatim overlap fragments (Winnowing) located in both paragraphs. Empty
   *  when none found or local-match disabled. Present whenever a copied
   *  substring of `kgramLength + windowSize − 1`+ chars survives inside an
   *  otherwise-dissimilar pair, even if sentence-level coverage is 0. */
  verbatimMatches: LocalMatch[];
}

/** Document-level comparison result. */
export interface DocumentComparison {
  /** Per-paragraph comparison details */
  paragraphs: ParagraphComparison[];
  /** Overall document coverage (weighted average of paragraph coverageA) */
  coverage: number;
}

/** Represents a duplicate paragraph match found within a single document. */
export interface DuplicateMatch {
  /** Index of the first occurrence */
  index: number;
  /** The paragraph text */
  text: string;
  /** Indices of all duplicate occurrences */
  duplicateIndices: number[];
  /** Similarity score for each duplicate (parallel to duplicateIndices) */
  similarityScores: number[];
  /** All verbatim overlap fragments (Winnowing) between this paragraph and its
   *  duplicates. Empty when none found or local-match disabled. */
  verbatimMatches: LocalMatch[];
}

// ---------------------------------------------------------------------------
// Local match (Winnowing) — public result types. The fingerprint and engine
// internals live in winnowing.ts; only the located-fragment results are public.
// ---------------------------------------------------------------------------

/** A located text fragment within a source document. */
export interface TextSpan {
  /** Paragraph index in the source document. */
  paragraphIndex: number;
  /** Char offset where the match starts within the paragraph. */
  start: number;
  /** Char offset where the match ends (exclusive). */
  end: number;
  /** The matched substring. */
  text: string;
}

/** A verbatim overlap between two documents (length >= minMatchLength). */
export interface LocalMatch {
  /** Match location in document 1. */
  fromDoc1: TextSpan;
  /** Match location in document 2. */
  fromDoc2: TextSpan;
  /** Length of the matched substring (characters). */
  length: number;
}
