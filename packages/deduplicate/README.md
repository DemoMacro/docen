# @docen/deduplicate

![npm version](https://img.shields.io/npm/v/@docen/deduplicate)
![npm downloads](https://img.shields.io/npm/dw/@docen/deduplicate)
![npm license](https://img.shields.io/npm/l/@docen/deduplicate)

> Document deduplication and similarity analysis utilities for Tiptap/ProseMirror JSON content.

## Features

- 🔍 **Duplicate Detection** - Find duplicate or similar paragraphs within documents
- 📊 **Similarity Calculation** - Calculate text similarity ratios (0-100%) using Levenshtein distance
- 🔗 **Cross-Document Comparison** - Compare two documents and find similar paragraphs
- 🎯 **Smart Matching** - Find the most similar text from a list of candidates
- 🌐 **Multilingual Support** - Works with English, Chinese, and other languages
- ⚙️ **Configurable Options** - Adjust similarity thresholds, whitespace handling, and case sensitivity
- 🚀 **High Performance** - Optimized algorithms for fast similarity calculations
- 🔒 **Full Type Safety** - Comprehensive TypeScript definitions for all functions

## Installation

```bash
# Install with npm
$ npm install @docen/deduplicate

# Install with yarn
$ yarn add @docen/deduplicate

# Install with pnpm
$ pnpm add @docen/deduplicate
```

## Quick Start

```typescript
import { findDuplicates } from "@docen/deduplicate";

// Your Tiptap/ProseMirror editor content
const document = {
  type: "doc",
  content: [
    {
      type: "paragraph",
      content: [{ type: "text", text: "机器学习是人工智能的一个重要分支。" }],
    },
    {
      type: "paragraph",
      content: [{ type: "text", text: "机器学习是人工智能的一个重要分支。" }],
    },
    {
      type: "paragraph",
      content: [{ type: "text", text: "深度学习是机器学习的子领域。" }],
    },
  ],
};

// Find duplicate paragraphs (85% similarity threshold)
const duplicates = findDuplicates(document, { threshold: 0.85 });

console.log(duplicates);
// Output:
// [
//   {
//     index: 0,
//     text: "机器学习是人工智能的一个重要分支。",
//     duplicates: [1],
//     similarities: [1.0]
//   }
// ]
```

## API Reference

### `extractParagraphs(doc)`

Extracts all paragraph text from a Tiptap JSON document.

**Parameters:**

- `doc: JSONContent` - Tiptap/ProseMirror document

**Returns:** `string[]` - Array of paragraph texts

```typescript
import { extractParagraphs } from "@docen/deduplicate";

const paragraphs = extractParagraphs(document);
// ["机器学习是人工智能的一个重要分支。", "深度学习是机器学习的子领域。"]
```

### `calculateSimilarity(text1, text2, options?)`

Calculates similarity ratio between two texts using Levenshtein distance.

**Parameters:**

- `text1: string` - First text
- `text2: string` - Second text
- `options?: DeduplicateOptions` - Configuration options

**Returns:** `number` - Similarity ratio between 0 (completely different) and 1 (identical)

```typescript
import { calculateSimilarity } from "@docen/deduplicate";

const similarity = calculateSimilarity(
  "机器学习是人工智能的一个重要分支。",
  "机器学习是人工智能的重要分支。",
  { ignoreCase: true, ignoreWhitespace: true },
);

console.log(similarity); // 0.94 (94% similar)
```

**Options:**

```typescript
interface DeduplicateOptions {
  threshold?: number; // Similarity threshold (0-1), default: 0.85
  ignoreWhitespace?: boolean; // Ignore whitespace differences, default: true
  ignoreCase?: boolean; // Ignore case differences, default: true
  filter?: (text: string) => boolean; // Filter function, default: () => true
}
```

### `findDuplicates(doc, options?)`

Finds duplicate/similar paragraphs in a document.

**Parameters:**

- `doc: JSONContent` - Tiptap/ProseMirror document
- `options?: DeduplicateOptions` - Configuration options

**Returns:** `DuplicateMatch[]` - Array of duplicate matches

```typescript
import { findDuplicates } from "@docen/deduplicate";

const duplicates = findDuplicates(document, {
  threshold: 0.85,
  ignoreWhitespace: true,
  ignoreCase: true,
});

// Result type:
interface DuplicateMatch {
  index: number; // Index of first occurrence
  text: string; // The paragraph text
  duplicates: number[]; // Indices of duplicate occurrences
  similarities: number[]; // Similarity scores for each duplicate
}
```

### `compareDocuments(doc1, doc2, options?)`

Compares two documents and finds similar paragraphs.

**Parameters:**

- `doc1: JSONContent` - First document
- `doc2: JSONContent` - Second document
- `options?: DeduplicateOptions` - Configuration options

**Returns:** `DocumentComparison[]` - Array of all paragraph comparisons

```typescript
import { compareDocuments } from "@docen/deduplicate";

const comparisons = compareDocuments(doc1, doc2, {
  filter: (text) => text.length >= 20, // Only compare texts >= 20 chars
});

// Result type:
interface DocumentComparison {
  fromDoc1: { index: number; text: string };
  fromDoc2: { index: number; text: string } | null; // null if no match
  similarity: number; // 0 if filtered or no match
  filtered: boolean; // true if filtered by filter function
}
```

**Note:** This function returns **all** paragraphs from `doc1`, including filtered ones. Use the `filtered` and `similarity` properties to determine the comparison result.

### `findMostSimilar(targetText, candidates, options?)`

Finds the most similar text from a list of candidates.

**Parameters:**

- `targetText: string` - Target text to match
- `candidates: string[]` - Array of candidate texts
- `options?: DeduplicateOptions` - Configuration options

**Returns:** `MostSimilarResult | null` - Best match or null if no candidates

```typescript
import { findMostSimilar } from "@docen/deduplicate";

const target = "人工智能的快速发展给各个行业带来了巨大的变化。";
const candidates = [
  "区块链技术不断发展和影响全球金融部门。",
  "人工智能的快速增长正在以显著的方式改变不同行业。",
  "气候变化仍然是人类面临的最紧迫的挑战之一。",
];

const result = findMostSimilar(target, candidates);

// Result:
// {
//   text: "人工智能的快速增长正在以显著的方式改变不同行业。",
//   index: 1,
//   similarity: 0.33
// }
```

### `distance(str1, str2)` & `closest(target, candidates)`

Calculate edit distance and find closest string matches.

```typescript
import { distance, closest } from "@docen/deduplicate";

// Calculate edit distance between two strings
const dist = distance("kitten", "sitting");
console.log(dist); // 3

// Find the closest string from candidates
const closestStr = closest("kitten", ["kitchen", "sitting", "kit"]);
console.log(closestStr); // "kitchen"
```

## Usage Examples

### Basic Duplicate Detection

```typescript
import { findDuplicates } from "@docen/deduplicate";

const document = {
  type: "doc",
  content: [
    { type: "paragraph", content: [{ type: "text", text: "First paragraph." }] },
    { type: "paragraph", content: [{ type: "text", text: "Duplicate paragraph." }] },
    { type: "paragraph", content: [{ type: "text", text: "Unique paragraph." }] },
    { type: "paragraph", content: [{ type: "text", text: "Duplicate paragraph." }] },
  ],
};

const duplicates = findDuplicates(document);

duplicates.forEach((dup) => {
  console.log(`Found "${dup.text}" at index ${dup.index}`);
  console.log(`  Duplicates at: ${dup.duplicates.join(", ")}`);
  console.log(
    `  Similarities: ${dup.similarities.map((s) => (s * 100).toFixed(1) + "%").join(", ")}`,
  );
});
```

### Document Comparison for Plagiarism Detection

```typescript
import { compareDocuments } from "@docen/deduplicate";

const studentEssay = parseEssay(studentSubmission);
const referenceEssay = parseEssay(referenceMaterial);

const comparisons = compareDocuments(studentEssay, referenceEssay, {
  filter: (text) => text.length >= 30, // Only compare substantial paragraphs
});

comparisons.forEach((comp) => {
  if (comp.similarity >= 0.75 && !comp.filtered) {
    console.log(`Suspicious similarity detected:`);
    console.log(`  Student: "${comp.fromDoc1.text}"`);
    console.log(`  Reference: "${comp.fromDoc2.text}"`);
    console.log(`  Similarity: ${(comp.similarity * 100).toFixed(1)}%`);
  }
});
```

### Custom Similarity Thresholds

```typescript
import { findDuplicates, calculateSimilarity } from "@docen/deduplicate";

// High precision (fewer false positives)
const exactDuplicates = findDuplicates(document, { threshold: 0.95 });

// High recall (catch more potential duplicates)
const looseMatches = findDuplicates(document, { threshold: 0.7 });

// Manual similarity calculation
const similarity = calculateSimilarity(
  "The quick brown fox jumps over the lazy dog.",
  "The quick brown cat jumps over the lazy dog.",
  { ignoreCase: true, ignoreWhitespace: true },
);

console.log(`Similarity: ${(similarity * 100).toFixed(1)}%`);
```

### Filtering Short Texts

```typescript
import { compareDocuments } from "@docen/deduplicate";

const comparisons = compareDocuments(doc1, doc2, {
  // Only compare paragraphs with at least 20 characters
  filter: (text) => text.length >= 20,
});

// Process results
comparisons.forEach((comp) => {
  if (comp.filtered) {
    console.log(`Skipped (too short): "${comp.fromDoc1.text}"`);
  } else if (comp.fromDoc2) {
    console.log(`Similarity: ${(comp.similarity * 100).toFixed(1)}%`);
    console.log(`  Doc1: "${comp.fromDoc1.text}"`);
    console.log(`  Doc2: "${comp.fromDoc2.text}"`);
  } else {
    console.log(`No match found: "${comp.fromDoc1.text}"`);
  }
});
```

### Advanced Filtering

```typescript
import { compareDocuments } from "@docen/deduplicate";

const comparisons = compareDocuments(doc1, doc2, {
  // Combine multiple conditions
  filter: (text) => {
    const minLength = 10;
    const maxLength = 500;
    const length = text.length;

    // Skip very short or very long texts
    if (length < minLength || length > maxLength) return false;

    // Skip numbered lists like "1." "2." etc.
    if (/^\d+\.$/.test(text.trim())) return false;

    // Skip URLs
    if (text.startsWith("http://") || text.startsWith("https://")) return false;

    return true;
  },
});
```

### Language-Specific Options

```typescript
import { calculateSimilarity } from "@docen/deduplicate";

// English: Case-insensitive comparison
const enSimilarity = calculateSimilarity("Hello World", "hello world", { ignoreCase: true });
// 1.0 (100% similar)

// Chinese: No case sensitivity needed
const zhSimilarity = calculateSimilarity(
  "机器学习是人工智能的重要分支",
  "机器学习是人工智能的重要分支",
  { ignoreCase: true },
);
// 1.0 (100% similar)

// Whitespace handling
const wsSimilarity = calculateSimilarity("机器学习　　是　一个　领域", "机器学习 是 一个 领域", {
  ignoreWhitespace: true,
});
// 1.0 (100% similar, full-width vs half-width spaces)
```

## Performance

Optimized algorithms for efficient document processing:

- **Time Complexity:** O(n×m) for text similarity calculation
- **Space Complexity:** O(min(n,m))
- **Scalability:** Handles large documents efficiently

For very large documents, consider processing in chunks or using Web Workers.

## TypeScript Types

All functions are fully typed with TypeScript:

```typescript
import type {
  DeduplicateOptions,
  DuplicateMatch,
  DocumentComparison,
  MostSimilarResult,
  JSONContent,
} from "@docen/deduplicate";
```

## Contributing

Contributions are welcome! Please read our [Contributor Covenant](https://www.contributor-covenant.org/version/2/1/code_of_conduct/) and submit pull requests to the [main repository](https://github.com/DemoMacro/docen).

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
