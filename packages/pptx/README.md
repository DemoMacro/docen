# @docen/pptx

![npm version](https://img.shields.io/npm/v/@docen/pptx)
![npm downloads](https://img.shields.io/npm/dw/@docen/pptx)
![npm license](https://img.shields.io/npm/l/@docen/pptx)

> The PPTX engine for Docen editors. Parse and generate presentations through `@office-open/pptx`, then project the canonical deck into drawing members for the shared LeaferJS canvas.

## Features

- 🎞️ **Canonical PPTX model** — read and write `PresentationOptions` without tying callers to the canvas
- 🧭 **Scene projection** — turn slides, shapes, pictures, tables, charts, media, and SmartArt into paint-ready drawing members
- 🎯 **Editable addressing** — hit-test, offset, resize, and rotate source children through stable scene paths
- 🧩 **Office-open types** — export the model and I/O types that PPTX editors need in one facade

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/pptx

# Install with npm
$ npm install @docen/pptx
```

## Quick Start

```typescript
import { generatePresentation, parsePresentation, projectPresentation } from "@docen/pptx";

const bytes = await generatePresentation({
  slides: [{ children: [{ shape: { x: 914400, y: 914400, width: 4572000, height: 2743200 } }] }],
});
const deck = await parsePresentation(bytes);
const scene = projectPresentation(deck);
```

Async APIs accept office-open's full `DataType`; `parsePresentationSync` accepts its byte-only `SyncDataType`. Generation defaults to a `Uint8Array`, with `OutputType` available for blobs, streams, Node buffers, and other output shapes.

## Architecture

```
@office-open/pptx   PresentationOptions ↔ PPTX
        ↓
@docen/pptx         canonical deck → projected drawing members
        ↓
@docen/core         LeaferJS scene painting
```

The projection owns slide geometry and source-child addressing. It does not own selection state, gestures, undo, or UI—those live in the editor host.

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
