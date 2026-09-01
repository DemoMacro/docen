# @docen/pretext

![npm version](https://img.shields.io/npm/v/@docen/pretext)
![npm downloads](https://img.shields.io/npm/dw/@docen/pretext)
![npm license](https://img.shields.io/npm/l/@docen/pretext)

> Vendored fork of [`@chenglou/pretext`](https://github.com/chenglou/pretext) 0.0.8 (commit `a79a6a5`) — fast, deterministic text measurement and line breaking — maintained in-tree so docen can carry Word/CJK layout fixes upstream won't.

The upstream project measures and breaks text through the canvas 2d API with pure arithmetic on top; the `rich-inline` entry adds multi-font inline runs, collapsed spaces, and extra-width atoms. This fork exists because docen's editor needs `edit == render` at Word-level CJK fidelity, which requires deep changes to the measurement internals — beyond what a patch file can sanely carry.

## Vendored changes

- **CJK canvas→DOM advance correction** (`measurement.ts`) — canvas `measureText` rounds fullwidth CJK advances up to whole pixels; at fractional font sizes (11pt = 14.67px) every glyph over-measures by ~0.33px, flipping wrapped line counts versus the DOM render. A per-font hidden-span probe measures the delta and `getCorrectedSegmentWidth` subtracts it per CJK grapheme.
- **Empty-text atom retention** (`rich-inline.ts`) — an inline item whose text is pure whitespace but carries `extraWidth` (an inline image's padding/border chrome) survives `prepareRichInline` as a zero-text unbreakable atom instead of being collapsed away.
- The `.` entry keeps the full upstream surface (`prepare` / `layout` / `layoutWithLines` …); docen consumes `./rich-inline` from [`@docen/layout`](../layout/README.md) (line breaking) and [`@docen/editor`](../editor/README.md) (paginator measurement).

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/pretext

# Install with npm
$ npm install @docen/pretext
```

## Quick Start

```typescript
import { prepareRichInline, layoutNextRichInlineLineRange } from "@docen/pretext/rich-inline";

const prepared = prepareRichInline([{ text: "甲乙丙丁", font: "16px serif" }]);
const range = layoutNextRichInlineLineRange(prepared, 64);
```

## License

- [MIT](./LICENSE) &copy; [Demo Macro](https://www.demomacro.com/) — includes vendored code &copy; Pretext contributors, MIT.
