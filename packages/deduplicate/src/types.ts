import type { JSONContent } from "@tiptap/core";

export type { JSONContent };

// ---------------------------------------------------------------------------
// Match classification
// ---------------------------------------------------------------------------

/** Classification of text similarity relationship. */
export type MatchKind = "contained" | "similar" | "weakOverlap" | "none";

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/** Sentence splitter function. */
export type SentenceSplitter = (text: string) => string[];

/** Configuration options for deduplication. */
export interface DeduplicateOptions {
  /** Minimum similarity threshold (0-1) for duplicate detection. @default 0.6 */
  threshold?: number;
  /** SimHash hamming distance threshold for candidate screening. @default 10 */
  hammingThreshold?: number;
  /** Levenshtein similarity threshold for fine-grained verification. @default 0.6 */
  levenshteinThreshold?: number;
  /** Minimum sentence length for SimHash fingerprinting and Phase 2 matching. @default 15 */
  minSentenceLength?: number;
  /** Custom sentence splitter. @default splitSentences (Chinese & English aware) */
  splitter?: SentenceSplitter;
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
}

// ---------------------------------------------------------------------------
// Result types
// ---------------------------------------------------------------------------

/** Bidirectional coverage between two paragraphs. */
export interface Coverage {
  /** Proportion of paragraph A's sentences matched in B (0-1) */
  covA: number;
  /** Proportion of paragraph B's sentences matched in A (0-1) */
  covB: number;
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
  /** Overall similarity score (max of covA, covB) */
  similarity: number;
}

/** Document-level comparison result. */
export interface DocumentResult {
  /** Per-paragraph comparison details */
  paragraphs: ParagraphComparison[];
  /** Overall document coverage (average of paragraph covA) */
  coverage: number;
}

/** Represents a duplicate paragraph match found within a single document. */
export interface DuplicateMatch {
  /** Index of the first occurrence */
  index: number;
  /** The paragraph text */
  text: string;
  /** Indices of all duplicate occurrences */
  duplicates: number[];
  /** Similarity scores for each duplicate */
  similarities: number[];
}
