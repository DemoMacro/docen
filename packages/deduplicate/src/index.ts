// Public API surface: 4 functions + supporting types. Implementation helpers
// (SimHash fingerprinting, sentence splitting, coverage classification, the
// Winnowing engine) stay internal — callers always go through compareDocuments
// or findDuplicates, with verbatim local-match folded in via the localMatch
// option rather than a separate entry point.
export type {
  JSONContent,
  MatchKind,
  SentenceSplitter,
  DeduplicateOptions,
  LocalMatchConfig,
  Coverage,
  ParagraphComparison,
  DocumentComparison,
  DuplicateMatch,
  LocalMatch,
  TextSpan,
} from "./types";

export {
  calculateSimilarity,
  compareDocuments,
  extractParagraphs,
  findDuplicates,
} from "./deduplicate";
