# docen

![npm version](https://img.shields.io/npm/v/docen)
![npm downloads](https://img.shields.io/npm/dw/docen)
![npm license](https://img.shields.io/npm/l/docen)

> Universal document toolkit — one package for headless Markdown/DOCX conversion AND the full `<docen-document>` web-component editor (via the `docen/editor` subpath).

## Features

- 🔄 **Universal Format Support** - Seamless conversion between Markdown and DOCX (styled HTML is accepted as paste input)
- 🎯 **Unified API** - Consistent, intuitive interface across all format conversions
- 📦 **All-in-One Package** - Single dependency for both headless conversion AND the full `<docen-document>` editor (via `docen/editor`)
- 🔧 **Built on TipTap** - Powered by the robust TipTap/ProseMirror ecosystem
- 💪 **TypeScript-First** - Full type safety with comprehensive TypeScript support
- ⚡ **Zero Configuration** - Works out of the box with smart defaults
- 🌳 **Extensible** - Customize conversions with TipTap extensions and advanced options
- 🔄 **Bidirectional** - Convert in both directions (parse ↔ generate) for each format

## Installation

```bash
# Install with npm
$ npm install docen

# Install with yarn
$ yarn add docen

# Install with pnpm
$ pnpm add docen
```

## Quick Start

### Markdown ↔ TipTap JSON

```typescript
import { parseMarkdown, generateMarkdown } from "docen";

// Parse Markdown to TipTap JSON
const doc = parseMarkdown("# Hello World\n\nThis is **bold** text.");

// Generate Markdown from TipTap JSON
const markdown = generateMarkdown(doc);
```

### DOCX ↔ TipTap JSON

```typescript
import { parseDOCX, generateDOCX } from "docen";

// Parse DOCX to TipTap JSON
const doc = parseDOCX(buffer);

// Generate DOCX from TipTap JSON
const docxBuffer = await generateDOCX(doc); // defaults to a Node.js Buffer
```

### Cross-Format Conversion

Convert between formats by using TipTap JSON as the intermediate format:

```typescript
import { parseMarkdown, generateDOCX } from "docen";

// Markdown → DOCX
const md = "# Title\n\nContent...";
const doc = parseMarkdown(md);
const docx = await generateDOCX(doc, { packer: { type: "blob" } });
```

### Full Editor (via `docen/editor`)

The same `docen` package also re-exports the turnkey web-component editor. Import from the `docen/editor` subpath to register `<docen-document>` and apply a theme:

```html
<docen-document id="doc" filename="Welcome.docx"></docen-document>

<script type="module">
  import { registerComponents, applyTheme } from "docen/editor";
  registerComponents();
  applyTheme("light");
</script>
```

> The editor lives on a subpath so that pure-converter imports (`import { parseDOCX } from "docen"`) stay tree-shakable and never pull in the Fluent UI shell.

## API Reference

### Markdown Functions

#### `parseMarkdown(markdown)`

Parses a Markdown string into TipTap JSON content.

**Parameters:**

- `markdown: string` - Markdown string to parse

**Returns:** `JSONContent` - TipTap document object

```typescript
const doc = parseMarkdown("# Hello\n\nWorld");
```

#### `generateMarkdown(doc)`

Generates a Markdown string from TipTap JSON content.

**Parameters:**

- `doc: JSONContent` - TipTap document object

**Returns:** `string` - Markdown string

```typescript
const markdown = generateMarkdown({ type: 'doc', content: [...] });
```

### DOCX Functions

#### `parseDOCX(input)`

Parses a DOCX file into TipTap JSON content.

**Parameters:**

- `input: Buffer | ArrayBuffer | Uint8Array | string` - DOCX file data or path

**Returns:** `JSONContent` - TipTap document object

```typescript
import { readFileSync } from "node:fs";
const buffer = readFileSync("document.docx");
const doc = parseDOCX(buffer);
```

#### `generateDOCX(docJson, options?)`

Generates a DOCX file from TipTap JSON asynchronously. Styling is derived from the TipTap attrs. By default runs `prepareDocument` first — fetching http image URLs and embedding them as data URLs (required: http images are otherwise dropped).

**Parameters:**

- `docJson: JSONContent` - TipTap document object
- `options?: DocxGenerateOptions` - `{ prepare?, packer? }`:
  - `prepare` (default `true`): `true` runs the default image pre-fetch; `false` skips it; a `PrepareStep[]` runs custom steps.
  - `packer`: `PackerOptions`; `type` controls the output format (`"nodebuffer"` default → Buffer, `"blob"`, `"arraybuffer"`, …).

**Returns:** `Promise<Buffer | Blob | ArrayBuffer | Uint8Array | string>` - DOCX data in the requested format

```typescript
// Default: prepare images, Node.js Buffer
const buffer = await generateDOCX(doc);

// Skip preparation, Browser Blob
const blob = await generateDOCX(doc, { prepare: false, packer: { type: "blob" } });
```

#### `generateDOCXSync(docJson, packerOptions?)`

Synchronous variant — fastest throughput, blocks the event loop. Does **not** run `prepareDocument` (it is async); call `await prepareDocument(doc)` first when http images need embedding.

**Returns:** `Buffer | Blob | ArrayBuffer | Uint8Array | string` - DOCX data in the requested format

```typescript
const buffer = generateDOCXSync(doc);
```

#### `generateDOCXStream(docJson, options?)`

Streams the DOCX as a `ReadableStream<Uint8Array>` — for large documents or HTTP responses. Runs `prepareDocument` by default (async).

**Returns:** `Promise<ReadableStream<Uint8Array>>`

```typescript
const stream = await generateDOCXStream(doc);
return new Response(stream);
```

## Advanced Usage

### Custom Extensions

Use custom TipTap extensions for Markdown conversions:

```typescript
import { CustomExtension } from "./custom-extension";
import { parseMarkdown, generateMarkdown } from "docen";

const doc = parseMarkdown(md, [CustomExtension]);
const mdContent = generateMarkdown(doc, [CustomExtension]);
```

### DOCX Template Patching

Replace `{{placeholders}}` in a DOCX template with TipTap-JSON content:

```typescript
import { patchDOCX, parseMarkdown } from "docen";

const result = await patchDOCX({
  template: templateBuffer,
  patches: {
    title: { content: parseMarkdown("# Report") },
    body: { content: parseMarkdown("## Section\n\nHello **world**.") },
  },
  outputType: "nodebuffer",
});
```

Each patch's `content` is compiled to DOCX (styling derived from attrs) and the first section's children replace the placeholder. `keepOriginalStyles`, `recursive`, and `placeholderDelimiters` mirror the underlying `@office-open/docx` `patchDocument`.

## Format Conversion Matrix

| From \ To    | Markdown | DOCX     |
| ------------ | -------- | -------- |
| **Markdown** | -        | via JSON |
| **DOCX**     | via JSON | -        |

All conversions go through TipTap JSON as the intermediate format, ensuring consistency and enabling cross-format transformations.

## Supported Content Types

### Text Formatting

- Bold, Italic, Underline, Strikethrough
- Superscript, Subscript
- Text highlights, colors, backgrounds
- Font families, sizes, line heights

### Block Elements

- Headings (H1-H6)
- Paragraphs with alignment
- Blockquotes
- Horizontal rules
- Code blocks with syntax highlighting

### Lists & Tables

- Bullet lists, ordered lists
- Task lists with checkboxes
- Tables with colspan/rowspan

### Media & Links

- Images with embedded base64
- Hyperlinks

## Use Cases

- **Content Management Systems** - Import/export documents in multiple formats
- **Documentation Tools** - Convert between Markdown and Word
- **Note-taking Apps** - Support various import/export formats
- **Report Generation** - Generate DOCX reports from Markdown templates
- **Content Migration** - Migrate content between different formats
- **Collaborative Editing** - Use TipTap editor with format support

## Under the Hood

`docen` ships three entry points: the **root** re-exports the high-level converters; **`docen/docx`** exposes the full engine (`createDocxEditor`, `docxExtensions`, resolve/compile/prepare, styles); **`docen/editor`** exposes the `<docen-document>` web component. It builds on:

- **@docen/docx** - DOCX / Markdown converters built on the DocxManager architecture (full surface via `docen/docx`)
- **@docen/editor** - Fluent UI shell + docx engine → `<docen-document>` (exposed via the `docen/editor` subpath)
- **@office-open/docx** - Native OOXML parse/generate (`parseDocument`, `generateDocument`, `patchDocument`)
- **@tiptap/markdown** - Markdown serialization (via @docen/docx)

## Comparison with Alternatives

| Feature     | docen | markdown-docx | mammoth | turndown |
| ----------- | ----- | ------------- | ------- | -------- |
| MD → DOCX   | ✅    | ✅            | ❌      | ❌       |
| DOCX → MD   | ✅    | ❌            | ❌      | ❌       |
| TypeScript  | ✅    | ✅            | ✅      | ✅       |
| Unified API | ✅    | ❌            | ❌      | ❌       |
| Extensible  | ✅    | ❌            | ❌      | ✅       |

## License

MIT © [Demo Macro](https://www.demomacro.com/)
