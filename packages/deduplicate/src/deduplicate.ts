import type {
  JSONContent,
  DeduplicateOptions,
  DuplicateMatch,
  DocumentComparison,
  MostSimilarResult,
} from "./types";
import { distance, closest } from "fastest-levenshtein";

/**
 * Normalizes text for comparison by removing extra whitespace and optionally converting case
 */
export function normalizeText(
  text: string,
  ignoreWhitespace: boolean,
  ignoreCase: boolean,
): string {
  let normalized = text;

  if (ignoreWhitespace) {
    // Replace multiple whitespace with single space and trim
    normalized = normalized.replace(/\s+/g, " ").trim();
  }

  if (ignoreCase) {
    normalized = normalized.toLowerCase();
  }

  return normalized;
}

/**
 * Extracts plain text from a node (including all marks and nested content)
 */
export function extractTextFromNode(node: JSONContent): string {
  if (!node) return "";

  if (node.type === "text" && node.text) {
    return node.text;
  }

  if (node.content && Array.isArray(node.content)) {
    return node.content.map(extractTextFromNode).join("");
  }

  return "";
}

/**
 * Extracts all paragraph text from Tiptap JSON document
 */
export function extractParagraphs(doc: JSONContent): string[] {
  const paragraphs: string[] = [];

  function traverse(node: JSONContent): void {
    if (!node) return;

    // Extract text from paragraph and heading nodes
    if (node.type === "paragraph" || node.type === "heading") {
      const text = extractTextFromNode(node);
      if (text && text.trim().length > 0) {
        paragraphs.push(text);
      }
    }

    // Recursively traverse content
    if (node.content && Array.isArray(node.content)) {
      for (const child of node.content) {
        traverse(child);
      }
    }
  }

  traverse(doc);
  return paragraphs;
}

/**
 * Calculates similarity ratio between two texts using Levenshtein distance
 * @returns number between 0 (completely different) and 1 (identical)
 */
export function calculateSimilarity(
  text1: string,
  text2: string,
  options: DeduplicateOptions = {},
): number {
  const { ignoreWhitespace = true, ignoreCase = true } = options;

  const normalized1 = normalizeText(text1, ignoreWhitespace, ignoreCase);
  const normalized2 = normalizeText(text2, ignoreWhitespace, ignoreCase);

  if (normalized1.length === 0 && normalized2.length === 0) {
    return 1; // Both empty, consider identical
  }

  if (normalized1.length === 0 || normalized2.length === 0) {
    return 0; // One empty, no similarity
  }

  const editDistance = distance(normalized1, normalized2);
  const maxLength = Math.max(normalized1.length, normalized2.length);

  // Convert distance to similarity ratio
  return 1 - editDistance / maxLength;
}

/**
 * Finds duplicate/similar paragraphs in a document
 */
export function findDuplicates(
  doc: JSONContent,
  options: DeduplicateOptions = {},
): DuplicateMatch[] {
  const { threshold = 0.85 } = options;

  const paragraphs = extractParagraphs(doc);
  const matches: DuplicateMatch[] = [];
  const processed = new Set<number>();

  for (let i = 0; i < paragraphs.length; i++) {
    if (processed.has(i)) continue;

    const current = paragraphs[i];
    const duplicateIndices: number[] = [];
    const similarities: number[] = [];

    for (let j = i + 1; j < paragraphs.length; j++) {
      if (processed.has(j)) continue;

      const similarity = calculateSimilarity(current, paragraphs[j], options);

      if (similarity >= threshold) {
        duplicateIndices.push(j);
        similarities.push(similarity);
        processed.add(j);
      }
    }

    if (duplicateIndices.length > 0) {
      matches.push({
        index: i,
        text: current,
        duplicates: duplicateIndices,
        similarities,
      });
      processed.add(i);
    }
  }

  return matches;
}

/**
 * Compares two documents and finds similar paragraphs
 * Returns all paragraph comparisons with their similarity scores
 * Filtered texts are included in results but marked with filtered: true
 */
export function compareDocuments(
  doc1: JSONContent,
  doc2: JSONContent,
  options: DeduplicateOptions = {},
): DocumentComparison[] {
  const { filter = () => true } = options;

  const paragraphs1 = extractParagraphs(doc1);
  const paragraphs2 = extractParagraphs(doc2);
  const comparisons: DocumentComparison[] = [];

  for (let i = 0; i < paragraphs1.length; i++) {
    const text1 = paragraphs1[i];

    // Check if text1 is filtered
    if (!filter(text1)) {
      comparisons.push({
        fromDoc1: { index: i, text: text1 },
        fromDoc2: null,
        similarity: 0,
        filtered: true,
      });
      continue;
    }

    // Find best match in doc2
    let bestMatch = -1;
    let bestSimilarity = 0;

    for (let j = 0; j < paragraphs2.length; j++) {
      const text2 = paragraphs2[j];

      // Skip filtered paragraphs in doc2
      if (!filter(text2)) continue;

      const similarity = calculateSimilarity(text1, text2, options);

      if (similarity > bestSimilarity) {
        bestSimilarity = similarity;
        bestMatch = j;
      }
    }

    if (bestMatch >= 0) {
      comparisons.push({
        fromDoc1: { index: i, text: text1 },
        fromDoc2: { index: bestMatch, text: paragraphs2[bestMatch] },
        similarity: bestSimilarity,
        filtered: false,
      });
    } else {
      // No match found (all candidates were filtered or no similar content)
      comparisons.push({
        fromDoc1: { index: i, text: text1 },
        fromDoc2: null,
        similarity: 0,
        filtered: false,
      });
    }
  }

  return comparisons;
}

/**
 * Finds the most similar paragraph from a list of candidates
 */
export function findMostSimilar(
  targetText: string,
  candidates: string[],
  options: DeduplicateOptions = {},
): MostSimilarResult | null {
  if (candidates.length === 0) return null;

  const { ignoreWhitespace = true, ignoreCase = true } = options;

  const normalizedTarget = normalizeText(targetText, ignoreWhitespace, ignoreCase);
  const normalizedCandidates = candidates.map((c) =>
    normalizeText(c, ignoreWhitespace, ignoreCase),
  );

  const closestMatch = closest(normalizedTarget, normalizedCandidates);
  const index = normalizedCandidates.indexOf(closestMatch);

  if (index === -1) return null;

  const similarity = calculateSimilarity(targetText, candidates[index], options);

  return {
    text: candidates[index],
    index,
    similarity,
  };
}

// Re-export fastest-levenshtein functions for convenience
export { distance, closest };
