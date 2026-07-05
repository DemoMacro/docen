# @docen/deduplicate

![npm version](https://img.shields.io/npm/v/@docen/deduplicate)
![npm downloads](https://img.shields.io/npm/dw/@docen/deduplicate)
![npm license](https://img.shields.io/npm/l/@docen/deduplicate)

> Document deduplication and similarity analysis for Tiptap/ProseMirror JSON content, using SimHash screening + Levenshtein verification.

## Features

- Duplicate detection within a single document
- Cross-document paragraph comparison with bidirectional coverage
- Sentence-level matching: SimHash for fast screening, Levenshtein for precise verification
- No false positives from n-gram containment — all matches verified by edit distance
- Multilingual support (Chinese, English, etc.)

## Installation

```bash
pnpm add @docen/deduplicate
```

## Quick Start

```typescript
import { findDuplicates } from "@docen/deduplicate";

const document = {
  type: "doc",
  content: [
    { type: "paragraph", content: [{ type: "text", text: "机器学习是人工智能的一个重要分支。" }] },
    { type: "paragraph", content: [{ type: "text", text: "机器学习是人工智能的一个重要分支。" }] },
    { type: "paragraph", content: [{ type: "text", text: "深度学习是机器学习的子领域。" }] },
  ],
};

const duplicates = findDuplicates(document, { similarityThreshold: 0.85 });
// [{ index: 0, text: "机器学习是...", duplicateIndices: [1], similarityScores: [1.0] }]
```

## API Reference

### `extractParagraphs(doc)`

Extracts all paragraph/heading text from a Tiptap JSON document. Consecutive paragraph nodes are merged when the first does not end with sentence-ending punctuation. DOCX text-box content (`wpsShape` / `wpgGroup`) is pulled out as standalone paragraphs — a text box's body never merges into its host paragraph.

```typescript
import { extractParagraphs } from "@docen/deduplicate";

const paragraphs = extractParagraphs(document);
// ["第一段。", "第二段。"]
```

### `calculateSimilarity(text1, text2)`

Calculates similarity using Levenshtein normalized distance.

```typescript
import { calculateSimilarity } from "@docen/deduplicate";

calculateSimilarity("你好世界", "你好世界"); // 1.0
calculateSimilarity("你好世界", "你好地球"); // ~0.5
calculateSimilarity("你好", "再见"); // ~0.0
```

### `findDuplicates(doc, options?)`

Finds duplicate/similar paragraphs within a single document.

```typescript
import { findDuplicates } from "@docen/deduplicate";

const duplicates = findDuplicates(document, {
  similarityThreshold: 0.85, // Minimum similarity (0-1), default: 0.6
});
```

### `compareDocuments(doc1, doc2, options?)`

Compares two documents and returns per-paragraph comparisons with bidirectional sentence-level coverage.

```typescript
import { compareDocuments } from "@docen/deduplicate";

const result = compareDocuments(doc1, doc2, {
  similarityThreshold: 0.6, // Noise floor below which → "none"
  hammingThreshold: 10, // SimHash screening distance
  levenshteinThreshold: 0.6, // Sentence-level verification threshold
});

result.paragraphs.forEach((pc) => {
  console.log(`[${pc.matchKind}] ${(pc.similarity * 100).toFixed(0)}%`);
  console.log(
    `  coverageA=${(pc.coverage.coverageA * 100).toFixed(0)}% coverageB=${(pc.coverage.coverageB * 100).toFixed(0)}%`,
  );
});
```

### Verbatim local match (built into `compareDocuments` / `findDuplicates`)

Both `compareDocuments` and `findDuplicates` also detect **verbatim copies** hidden inside dissimilar text — the "a hundred-char paragraph with a dozen copied characters" case that whole-paragraph SimHash dilutes and overall Levenshtein misses. It runs automatically via the `localMatch` option (on by default) and the fragments land in each result's `verbatimMatches` field:

```typescript
import { compareDocuments } from "@docen/deduplicate";

const result = compareDocuments(doc1, doc2);
for (const pc of result.paragraphs) {
  for (const m of pc.verbatimMatches) {
    console.log(`copied fragment: "${m.fromDoc1.text}" (${m.length} chars)`);
    // m.fromDoc1 / m.fromDoc2 each carry { paragraphIndex, start, end, text } for highlighting
  }
}
```

Pass `localMatch: false` to disable, or `{ kgramLength, windowSize, minMatchLength }` to tune. The default `kgramLength=10, windowSize=6` gives a 15-char guarantee (`t = kgramLength + windowSize − 1`): any shared substring of 15+ characters within a paragraph pair is reported.

## Options

```typescript
interface DeduplicateOptions {
  /** Minimum similarity (0-1). @default 0.6 */
  similarityThreshold?: number;
  /** SimHash hamming distance for candidate screening. @default 10 */
  hammingThreshold?: number;
  /** Levenshtein similarity for sentence verification. @default 0.6 */
  levenshteinThreshold?: number;
  /** Minimum sentence length for SimHash fingerprinting. @default 15 */
  minSentenceLength?: number;
  /** Custom sentence splitter (Chinese & English aware by default). */
  splitter?: (text: string) => string[];
  /** Verbatim local-match (Winnowing). `false` disables; an object tunes
   *  { kgramLength, windowSize, minMatchLength }. @default enabled (k=10, w=6 ⇒ 15-char guarantee) */
  localMatch?: boolean | LocalMatchConfig;
}
```

## Result Types

```typescript
interface DocumentComparison {
  paragraphs: ParagraphComparison[];
  coverage: number; // Average of paragraph coverageA
}

interface ParagraphComparison {
  fromDoc1: { index: number; text: string };
  fromDoc2: { index: number; text: string } | null;
  coverage: { coverageA: number; coverageB: number };
  matchKind: "contained" | "similar" | "partial" | "none";
  similarity: number; // max(coverageA, coverageB)
  verbatimMatches: LocalMatch[]; // verbatim fragments (Winnowing)
}

interface DuplicateMatch {
  index: number;
  text: string;
  duplicateIndices: number[];
  similarityScores: number[];
  verbatimMatches: LocalMatch[]; // verbatim fragments vs all duplicates
}

interface LocalMatch {
  fromDoc1: TextSpan; // { paragraphIndex, start, end, text }
  fromDoc2: TextSpan;
  length: number; // matched characters
}
```

### Match Classification

| Kind        | Condition                                        | Meaning                                     |
| ----------- | ------------------------------------------------ | ------------------------------------------- |
| `contained` | max(coverageA, coverageB) >= 0.8                 | One paragraph mostly contained in the other |
| `similar`   | min(coverageA, coverageB) >= 0.6                 | High bidirectional overlap                  |
| `partial`   | max(coverageA, coverageB) >= similarityThreshold | Partial overlap                             |
| `none`      | max(coverageA, coverageB) < similarityThreshold  | No meaningful match                         |

## How It Works

1. **Extract paragraphs** from Tiptap JSON, split into sentences
2. **SimHash fingerprinting** — each sentence >= `minSentenceLength` gets a fingerprint, and each paragraph gets a paragraph-level fingerprint
3. **Paragraph-pair prescreen** — `hammingDistance` on paragraph fingerprints skips unlikely pairs before the expensive sentence matching (short paragraphs without a fingerprint bypass prescreening)
4. **Two-phase sentence matching** for candidate pairs:
   - Phase 1: SimHash hamming distance screens sentences (fast)
   - Phase 2: Levenshtein normalized similarity verifies matches (precise)
   - Unmatched short sentences: direct Levenshtein comparison
5. **No containment fallback** — eliminates false positives from n-gram coincidence
6. **Noise floor** controlled by `similarityThreshold` — matches below this are classified as "none"

7. **Verbatim local match (Winnowing)** — alongside sentence matching, each candidate pair runs k-gram fingerprints (windowed-min selection) matched by hash; a collision seeds a char-by-char `extendSeed` walk that recovers the full copied fragment of any length. Fragments ≥ `minMatchLength` land in `verbatimMatches` and upgrade an otherwise-`none` pair to `partial`. This is built into `compareDocuments` / `findDuplicates` — no separate function to choose. Guarantee (Schleimer et al. 2003): any shared substring of `t = kgramLength + windowSize − 1` chars yields ≥1 fragment. Built on `@nlptools/distance`'s `ngrams` + `fnv1a`; only the windowed-min selection and seed-extend are docen's own.

## License

MIT &copy; [Demo Macro](https://www.demomacro.com/)
