# @docen/core

![npm version](https://img.shields.io/npm/v/@docen/core)
![npm downloads](https://img.shields.io/npm/dw/@docen/core)
![npm license](https://img.shields.io/npm/l/@docen/core)

> The scene painter for the docen editors: it instantiates a laid-out `LayoutDoc` ([@docen/layout](../layout/README.md)) as a LeaferJS tree.

> Consumed by [`@docen/editor`](../editor/README.md)'s canvas stage, one paint per page sync. The future pptx/xlsx canvas editors reuse the same painter.

## Features

- 🎨 **Painter** — `paintScene(tree, items, ctx)` walks a laid-out page and builds the LeaferJS tree: text lines, tables, drawing members (pictures / shapes / paths / text boxes, incl. vector metafile members), and page furniture
- 🧱 **Dumb by design** — the layout engine owns ALL geometry; the painter positions what it is given and never measures. Text elements carry explicit width AND height (Leafer never paints an element whose height is still 0)
- 🖼️ **Drawing hit boxes** — every drawing paints a page-local hit box so clicks resolve back to the editing model's drawing selection
- 🧩 **Headless-ready** — no DOM assumptions beyond LeaferJS itself; runs under `@leafer-ui/node` for export / thumbnails

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/core

# Install with npm
$ npm install @docen/core
```

## Quick Start

```typescript
import { paintScene, type PaintContext } from "@docen/core";
import { App, type IGroup } from "leafer-ui";

// items: the laid-out page from @docen/layout (stackBlocks/TextMeasurer),
// ctx: the page's flow box + furniture + background (the docx projection).
const app = new App({ view, editor: {} });
const tree = app.tree as IGroup;
const ctx: PaintContext = {/* flow, furniture, background, hitBoxes... */};
paintScene(tree, items, ctx);
```

## Architecture

```
@docen/layout      LayoutDoc — paginated geometry (engine owns all measurement)
        ↓
@docen/core        this package — LayoutDoc → LeaferJS elements
        ↓
@docen/editor      canvas stage (App lifecycle, zoom, caret/selection mapping)
```

The painter deliberately owns **no layout decisions** (geometry comes from the engine) and **no editing semantics** (selection, undo, keyboard/IME live in the editor).

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
