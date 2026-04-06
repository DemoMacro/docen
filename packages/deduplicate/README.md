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

const duplicates = findDuplicates(document, { threshold: 0.85 });
// [{ index: 0, text: "机器学习是...", duplicates: [1], similarities: [1.0] }]
```

## API Reference

### `extractParagraphs(doc)`

Extracts all paragraph/heading text from a Tiptap JSON document. Consecutive paragraph nodes are merged when the first does not end with sentence-ending punctuation.

```typescript
import { extractParagraphs } from "@docen/deduplicate";

const paragraphs = extractParagraphs(document);
// ["第一段。", "第二段。"]
```

### `splitSentences(text)`

Splits text into sentences. Chinese-aware (supports 。！？；.!?;).

```typescript
import { splitSentences } from "@docen/deduplicate";

const sentences = splitSentences("第一句。第二句！第三句？");
// ["第一句。", "第二句！", "第三句？"]
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
  threshold: 0.85, // Minimum similarity (0-1), default: 0.6
});
```

### `compareDocuments(doc1, doc2, options?)`

Compares two documents and returns per-paragraph comparisons with bidirectional sentence-level coverage.

```typescript
import { compareDocuments } from "@docen/deduplicate";

const result = compareDocuments(doc1, doc2, {
  threshold: 0.6, // Noise floor below which → "none"
  hammingThreshold: 10, // SimHash screening distance
  levenshteinThreshold: 0.6, // Sentence-level verification threshold
});

result.paragraphs.forEach((pc) => {
  console.log(`[${pc.matchKind}] ${(pc.similarity * 100).toFixed(0)}%`);
  console.log(
    `  covA=${(pc.coverage.covA * 100).toFixed(0)}% covB=${(pc.coverage.covB * 100).toFixed(0)}%`,
  );
});
```

### `findBestMatch` (re-export from `@nlptools/distance`)

One-shot fuzzy search: find the best matching string from candidates.

```typescript
import { findBestMatch } from "@docen/deduplicate";

const result = findBestMatch("kitten", ["sitting", "kit", "mitten"]);
// { item: "kit", score: 0.5, index: 1 }
```

## Options

```typescript
interface DeduplicateOptions {
  /** Minimum similarity threshold (0-1). @default 0.6 */
  threshold?: number;
  /** SimHash hamming distance for candidate screening. @default 10 */
  hammingThreshold?: number;
  /** Levenshtein similarity for sentence verification. @default 0.6 */
  levenshteinThreshold?: number;
  /** Minimum sentence length for SimHash fingerprinting. @default 15 */
  minSentenceLength?: number;
  /** Custom sentence splitter (Chinese & English aware by default). */
  splitter?: (text: string) => string[];
}
```

## Result Types

```typescript
interface DocumentResult {
  paragraphs: ParagraphComparison[];
  coverage: number; // Average of paragraph covA
}

interface ParagraphComparison {
  fromDoc1: { index: number; text: string };
  fromDoc2: { index: number; text: string } | null;
  coverage: { covA: number; covB: number };
  matchKind: "contained" | "similar" | "weakOverlap" | "none";
  similarity: number; // max(covA, covB)
}

interface DuplicateMatch {
  index: number;
  text: string;
  duplicates: number[];
  similarities: number[];
}
```

### Match Classification

| Kind          | Condition                    | Meaning                                     |
| ------------- | ---------------------------- | ------------------------------------------- |
| `contained`   | max(covA, covB) >= 0.8       | One paragraph mostly contained in the other |
| `similar`     | min(covA, covB) >= 0.6       | High bidirectional overlap                  |
| `weakOverlap` | max(covA, covB) >= threshold | Partial overlap (only when threshold < 0.6) |
| `none`        | max(covA, covB) < threshold  | No meaningful match                         |

## How It Works

1. **Extract paragraphs** from Tiptap JSON, split into sentences
2. **SimHash fingerprinting** for sentences >= `minSentenceLength` characters
3. **Two-phase matching** per paragraph pair:
   - Phase 1: SimHash hamming distance screens candidates (fast)
   - Phase 2: Levenshtein normalized similarity verifies matches (precise)
   - Unmatched short sentences: direct Levenshtein comparison
4. **No containment fallback** — eliminates false positives from n-gram coincidence
5. **Noise floor** controlled by `threshold` — matches below this are classified as "none"

## License

MIT &copy; [Demo Macro](https://imst.xyz/)
