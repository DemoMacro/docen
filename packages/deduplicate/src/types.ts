import type { JSONContent } from "@tiptap/core";

/**
 * Configuration options for deduplication
 */
export interface DeduplicateOptions {
  /**
   * Minimum similarity threshold (0-1).
   * Paragraphs with similarity >= threshold will be considered duplicates.
   * @default 0.85
   */
  threshold?: number;
  /**
   * Whether to ignore whitespace differences when comparing
   * @default true
   */
  ignoreWhitespace?: boolean;
  /**
   * Whether to ignore case differences when comparing
   * @default true
   */
  ignoreCase?: boolean;
}

/**
 * Represents a duplicate paragraph match found in a document
 */
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

/**
 * Represents a comparison result between two documents
 */
export interface DocumentComparison {
  /** Similar paragraph from document 1 */
  fromDoc1: { index: number; text: string };
  /** Similar paragraph from document 2 */
  fromDoc2: { index: number; text: string };
  /** Similarity score for the pair */
  similarity: number;
}

/**
 * Result of finding the most similar text
 */
export interface MostSimilarResult {
  /** The most similar text */
  text: string;
  /** Index in candidates array */
  index: number;
  /** Similarity score */
  similarity: number;
}

// Re-export JSONContent for convenience
export type { JSONContent };
