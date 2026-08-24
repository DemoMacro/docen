# @docen/layout

Pure TypeScript layout engine for docen — zero DOM, zero rendering libraries. The inner ring (`text/`, `block/`) measures and breaks text and lays out paragraphs/tables; the outer ring builds on it (`flow/` — docx page boxing — is in; `fixed/` for pptx shape text and `grid/` for xlsx come with their editors).

## Design

- **Input is a `LayoutDoc` projection, all-px, style-cascades already resolved.** Adapters (docx from Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) convert their document units and resolve their style chains exactly once; the engine never sees a `styleId`, a twip, or a Tiptap node. Only layout semantics keep their OOXML shape (line rules, grid pitch, snap flags) — those are the rules this engine exists to implement.
- **FontBackend is the measuring seam.** The engine asks a backend for a face and works from per-code-unit advances (opentype.js today; bundled pack / `queryLocalFonts` / Node `fs` backends land with the editors). Misses fall back with a warning — Word's behavior.
- **One packer for text, hard breaks, and inline pictures** — the unified breaker the DOM route never had. UAX #14 line breaking (linebreak.js) with CJK kinsoku, trailing-space hanging, first-line indent, float-zone width reduction, and per-line OOXML line-height semantics (exact / atLeast / multiple × docGrid pitch, CJK ceil snap).
- **Determinism is a contract.** Same input → same output every pass — the property the paginator's convergence depends on.

Known boundaries: no GSUB shaping (ligature substitution — advances run slightly wide, conservative for breaking); rowspan cell content counts fully on its start row; punctuation compression is not yet implemented; `flow/` does not yet thread float zones (page-level float narrowing lands with the renderer).

## Usage

```ts
import { OpentypeFontBackend, TextMeasurer, layoutBlock } from "@docen/layout";

const backend = new OpentypeFontBackend();
backend.register("Inter", fontBytes);

const measurer = new TextMeasurer(backend);
const laid = layoutBlock(paragraph, 612, { linePitchPx: 25 }, measurer);
// laid.lines — y, height, positioned items, split points
```
