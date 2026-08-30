# leafer-x-metafile

> Zero-dependency WMF / EMF+ metafile player: replays placeable-WMF record streams and EMF+ (GDI+) record streams into neutral, box-local drawing members — pictures, shapes, paths, and text boxes — that any renderer can consume.

Browsers cannot rasterize GDI metafiles, so Office documents carrying `.wmf`/`.emf` vector art render as empty frames. This package reassembles those byte streams (including dual-mode files whose real art hides in `META_ESCAPE` "WMFC" chunks as an embedded EMF+ stream) and replays them record-by-record into structured members, ready to be painted on canvas or instantiated as LeaferJS elements.

## Features

- 🎞️ **WMF replay** — `wmfMembers` walks a placeable-WMF stream (pens, brushes, fonts, DC state, polygons, polylines, rectangles, ellipses, `ExtTextOut` with per-byte Dx tracking, `StretchDIBits` blts)
- 🖼️ **EMF+ replay** — `emfPlusMembers` reassembles the WMFC-carried EMF and replays its GDI+ records (fill/draw paths, draw-image points, world transforms, `DrawString`), layered with the GDI-side text the exporters keep there
- 🧱 **DIB fallback** — `wmfDibFallback` / `bmpDataUrl` re-frame the largest blt-record DIB as a BMP data URL when no vector replay applies
- 📦 **Neutral output** — members are plain data (`MetafileMember`), geometry in box-local px, text as runs; no Leafer, no DOM, no dependencies

## Installation

```bash
# Install with pnpm
$ pnpm add leafer-x-metafile

# Install with npm
$ npm install leafer-x-metafile
```

## Quick Start

```typescript
import { emfPlusMembers, wmfMembers } from "leafer-x-metafile";

const bytes = new Uint8Array(/* the .wmf/.emf file */);
const members = emfPlusMembers(bytes, 400, 300) ?? wmfMembers(bytes, 400, 300);

for (const member of members ?? []) {
  if (member.kind === "path") drawPath(member.d, member.fill, member.line);
}
```

## Roadmap

- A LeaferJS painter layer that instantiates the members directly into a scene (`Leafer` / `Path` / `Text` nodes) — planned as the package's rendering half.

## License

[MIT](./LICENSE)
