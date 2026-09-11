# @docen/markdown

![npm version](https://img.shields.io/npm/v/@docen/markdown)
![npm downloads](https://img.shields.io/npm/dw/@docen/markdown)
![npm license](https://img.shields.io/npm/l/@docen/markdown)

> Format-agnostic Markdown syntax layer for docen — parses Markdown into a neutral document IR and renders it back. Format packages bind the IR to their own document models through a `MarkdownMapper`.

## Overview

`@docen/markdown` owns only the Markdown syntax itself (CommonMark + GFM tables): heading levels, nested lists, tables, quotes, code fences, and inline emphasis. It knows nothing about DOCX/PPTX/XLSX — each format package ships its own mapper:

- [`@docen/docx`](https://www.npmjs.com/package/@docen/docx) — the reference mapper, re-exported as the familiar one-argument `parseMarkdown` / `generateMarkdown`.

A future `@docen/pptx` / `@docen/xlsx` would follow the same pattern: implement `MarkdownMapper<T>` and bind the two generic functions.

## Installation

```bash
$ pnpm add @docen/markdown
```

## Usage

```typescript
import { parseMarkdown, generateMarkdown, type MarkdownMapper } from "@docen/markdown";

// Your format's document model — the mapper is the only contract.
const mapper: MarkdownMapper<MyNode> = {
  parseBlock: (block) => myParseBlock(block),
  parseInline: (inlines) => myParseInline(inlines),
  toBlocks: (doc) => myToBlocks(doc),
};

const nodes = parseMarkdown("# Title\n\nHello", mapper);
const markdown = generateMarkdown(nodes, mapper);
```

## Known gaps

- GFM footnote definitions (`[^1]: note`) and task-list checkboxes (`- [x]`) round-trip through the IR, but a format mapper decides what they mean — `@docen/docx` renders footnotes as plain `[^label]: text` paragraphs and task items as bullets with a literal checkbox prefix.
- Code-fence info strings ride the IR (`lang`); whether a format can store one is up to its mapper (DOCX has no code-language field, so a DOCX round-trip drops it).
- Inline HTML degrades to literal text; block-level HTML keeps its raw markup as a text paragraph (no HTML output pipeline exists).

## License

MIT — see [LICENSE](../../LICENSE).
