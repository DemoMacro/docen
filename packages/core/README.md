# @docen/core

![npm version](https://img.shields.io/npm/v/@docen/core)
![npm downloads](https://img.shields.io/npm/dw/@docen/core)
![npm license](https://img.shields.io/npm/l/@docen/core)

> Pure rendering layer mapping OOXML drawing data (image / shape / chart / smartart) to LeaferJS / ECharts elements, shared across the docx (Tiptap), pptx (LeaferJS), and xlsx (RevoGrid) editors.

> Consumed by [`@docen/editor`](../editor/README.md), which mounts these elements in `<docen-image>` web components with selection, resize, and rotate UX.

## Features

- 🖼️ **Image** — `renderImage` / `parseImage` map OOXML `<pic:pic>` (blip + transform + srcRect crop + outline) to LeaferJS `Image` elements
- 📐 **Geometry** — EMU ↔ px, srcRect → clip, transform math as stateless functions (reuse `@office-open/core`'s `convertEmuToPixels`)
- 🎨 **Style** — `renderFill` / `renderOutline` map OOXML `FillOptions` / `OutlineOptions` to LeaferJS paints
- 📤 **Export** — `exportImage` / `exportCanvas` wrap LeaferJS export (canvas → png/jpg base64) for DOCX `<a:blip>` round-trip
- 🧩 **Headless-ready** — DOM-free functions run identically in the browser and in Node (`@leafer-ui/node`) for SSR / thumbnails

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/core

# Install with npm
$ npm install @docen/core
```

## Quick Start

```typescript
import { renderImage, parseImage } from "@docen/core/image";
import { exportImage } from "@docen/core/export";
import { App, Image } from "leafer-ui";

// Render an OOXML image to LeaferJS
const options = renderImage({
  src: "data:image/png;base64,...",
  width: 400,
  height: 300,
  rotation: 15,
  crop: { left: 10000, top: 5000 }, // permyriad (0–100000 per side)
});

const app = new App({ view, editor: {} });
const image = new Image(options);
app.add(image);

// Read back the edited geometry after a user resize/rotate
const { width, height, rotation } = parseImage(image);

// Export the canvas to a base64 data URL for DOCX round-trip
const dataUrl = await exportImage(app, "png");
```

## Architecture

```
@office-open/core   OOXML data model (read / write .docx .pptx .xlsx)
        ↓
@docen/core         OOXML data ↔ visual elements  (this package)
        ↓
@docen/editor       UI components + NodeViews + interaction
```

`@docen/core` is the visual counterpart of `@office-open/core`'s data model. It owns the `render*` (data → element) and `parse*` (element → data) mappings plus geometry math and PNG/SVG export. It deliberately owns **no editing semantics**.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
