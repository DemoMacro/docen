# @docen/docx

![npm version](https://img.shields.io/npm/v/@docen/docx)
![npm downloads](https://img.shields.io/npm/dw/@docen/docx)
![npm license](https://img.shields.io/npm/l/@docen/docx)

> DOCX editor and converter powered by @office-open/docx with Tiptap editing layer, supporting bidirectional conversion between DOCX and Markdown (styled HTML is accepted as paste input).

> Need a ready-made visual editor? [`@docen/editor`](../editor/README.md) wraps this engine in a Fluent UI host with the turnkey `<docen-document>` web component.

## Features

- 📝 **Viewless Tiptap engine** — the DOCX editing model with DOCX-aware extensions; rendering belongs to the host (e.g. @docen/editor's canvas stage)
- 🔄 **DOCX Round-trip** — Near-lossless DOCX ↔ Editor conversion via @office-open/docx
- 📄 **Markdown Support** — Tiptap JSON ↔ Markdown conversion
- 🎨 **DOCX Properties** — Custom Tiptap extensions carry shading, borders, indent, spacing, floating, crop
- 🔗 **Template Patching** — Replace `{{placeholders}}` in DOCX templates with Tiptap-JSON content

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/docx

# Install with npm
$ npm install @docen/docx
```

## Quick Start

```typescript
import { docxExtensions, parseDOCX, generateDOCX } from "@docen/docx";
import { Editor } from "@docen/docx/core";

// Viewless: the editor is the editing model — rendering belongs to the host.
const editor = new Editor({
  element: null,
  extensions: docxExtensions,
  content: await parseDOCX(buffer),
});

// Save back to DOCX
const output = await generateDOCX(editor.getJSON());
```

## API

### Standalone Functions (Core)

These work without an editor instance — for headless/server/batch use.

```typescript
import {
  parseDOCX,
  generateDOCX,
  generateDOCXSync,
  generateDOCXStream,
  parseMarkdown,
  generateMarkdown,
} from "@docen/docx";

// DOCX pipeline: DOCX binary ↔ Tiptap JSON
const json = parseDOCX(buffer); // → JSONContent
const buffer = await generateDOCX(json); // → Buffer (pre-fetches http images by default)
const blob = await generateDOCX(json, { packer: { type: "blob" } }); // → Blob
const sync = generateDOCXSync(json); // → Buffer (skips prepare)
const stream = await generateDOCXStream(json); // → ReadableStream<Uint8Array>

// Markdown pipeline: Markdown string ↔ Tiptap JSON
const json = parseMarkdown("# Hello"); // → JSONContent
const md = generateMarkdown(json); // → string
```

### Editor

```typescript
import { createDocxEditor, docxExtensions } from "@docen/docx";

const editor = createDocxEditor({
  element: document.querySelector("#editor"),
  extensions: docxExtensions,
  spellcheck: true,
  editable: true,
});
```

### Extension Commands (Thin Wrappers)

Convenience commands that call standalone functions internally.

```typescript
// Load DOCX into editor (calls parseDOCX → setContent)
editor.commands.importDocx(buffer);

// Export editor content as DOCX (calls getJSON → generateDOCX)
const buffer = await editor.commands.exportDocx();
```

### Template Patching

Replace `{{placeholders}}` in a DOCX template with Tiptap-JSON content. Each
patch's `content` is prepared (default: fetch http images) then compiled to DOCX.

```typescript
import { patchDOCX, parseMarkdown } from "@docen/docx";

const result = await patchDOCX({
  template: templateBuffer,
  patches: {
    title: { content: parseMarkdown("# Report") },
  },
  outputType: "nodebuffer",
});
```

### Advanced: Model Bridge

`generateDOCX` runs `prepareDocument → compileDocument → generateDocument`
internally. You rarely need these directly — reach for them only when working
with the intermediate `DocumentOptions` (the OOXML persistence model):

```typescript
import { resolveDocument, compileDocument, prepareDocument } from "@docen/docx";

const json = resolveDocument(docOpts); // DocumentOptions → JSONContent
const docOpts = compileDocument(json); // JSONContent → DocumentOptions
await prepareDocument(json); // in place: http image URLs → data URLs
```

## Architecture

```
Standalone Functions (core)
  parseDOCX / generateDOCX / generateDOCXSync / generateDOCXStream / patchDOCX
  parseMarkdown / generateMarkdown
  resolveDocument / compileDocument / prepareDocument  (model bridge, advanced)
        ↕ used by
Tiptap Extension Commands (thin wrappers)
  editor.commands.importDocx() / exportDocx()
```

- **Runtime model**: Tiptap JSON with DOCX-rich attributes via custom extensions
- **Persistence model**: DocumentOptions (complete OOXML expressiveness)
- **Standalone functions are core** — extension commands are thin wrappers

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
