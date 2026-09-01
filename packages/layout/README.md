# @docen/layout

![npm version](https://img.shields.io/npm/v/@docen/layout)
![npm downloads](https://img.shields.io/npm/dw/@docen/layout)
![npm license](https://img.shields.io/npm/l/@docen/layout)

> Pure TypeScript layout engine for docen — zero DOM, zero rendering libraries. The inner ring (`text/`, `block/`) measures and breaks text and lays out paragraphs/tables; the outer ring builds on it (`flow/` — docx page boxing — is in; `fixed/` for pptx shape text and `grid/` for xlsx come with their editors).

Consumed by [`@docen/docx`](../docx/README.md)'s projection (DocumentOptions → LayoutDoc) and painted by [`@docen/core`](../core/README.md)'s canvas stage.

## Design

- **Input is a `LayoutDoc` projection, all-px, style-cascades already resolved.** Adapters (docx from Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) convert their document units and resolve their style chains exactly once; the engine never sees a `styleId`, a twip, or a Tiptap node. Only layout semantics keep their OOXML shape (line rules, grid pitch, snap flags) — those are the rules this engine exists to implement.
- **FontMetrics is the measuring seam.** The engine works from per-code-unit advances supplied by a `FontMetrics` provider (`browserFontMetrics` measures through canvas 2d today; headless and server providers slot into the same interface). Misses fall back with a warning — Word's behavior.
- **One packer for text, hard breaks, and inline pictures** — the unified breaker the DOM route never had. UAX #14 line breaking (linebreak.js) with CJK kinsoku, trailing-space hanging, first-line indent, float-zone width reduction, and per-line OOXML line-height semantics (exact / atLeast / multiple × docGrid pitch, CJK ceil snap).
- **Determinism is a contract.** Same input → same output every pass — the property the paginator's convergence depends on.

Known boundaries: no GSUB shaping (ligature substitution — advances run slightly wide, conservative for breaking); rowspan cell content counts fully on its start row.

## Usage

```ts
import { TextMeasurer, browserFontMetrics, layoutBlock } from "@docen/layout";

const measurer = new TextMeasurer(browserFontMetrics);
const laid = layoutBlock(paragraph, 612, { linePitchPx: 25 }, measurer);
// laid.lines — y, height, positioned items, split points
```

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
