# docen

![npm version](https://img.shields.io/npm/v/docen)
![npm downloads](https://img.shields.io/npm/dw/docen)
![npm license](https://img.shields.io/npm/l/docen)

> Universal document format converter providing unified API for seamless transformation between Markdown, HTML, and DOCX formats.

## Features

- 🔄 **Universal Format Support** - Seamless conversion between Markdown, HTML, and DOCX
- 🎯 **Unified API** - Consistent, intuitive interface across all format conversions
- 📦 **All-in-One Package** - Single dependency for all your document conversion needs
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

### HTML ↔ TipTap JSON

```typescript
import { parseHTML, generateHTML } from "docen";

// Parse HTML to TipTap JSON
const doc = parseHTML("<h1>Hello World</h1><p>This is <strong>bold</strong> text.</p>");

// Generate HTML from TipTap JSON
const html = generateHTML(doc);
```

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

Convert between any formats by using TipTap JSON as the intermediate format:

```typescript
import { parseHTML, generateDOCX } from "docen";

// HTML → DOCX
const html = "<h1>Title</h1><p>Content...</p>";
const doc = parseHTML(html);
const docx = await generateDOCX(doc, { packer: { type: "blob" } });

// Markdown → HTML
import { parseMarkdown, generateHTML } from "docen";
const md = "# Title\n\nContent...";
const doc2 = parseMarkdown(md);
const htmlContent = generateHTML(doc2);
```

## API Reference

### HTML Functions

#### `parseHTML(html, extensions?, options?)`

Parses an HTML string into TipTap JSON content.

**Parameters:**

- `html: string` - HTML string to parse
- `extensions?: Extensions` - Optional TipTap extensions (defaults to @docen/docx's docxExtensions)
- `options?: ParseOptions` - Optional ProseMirror parse options

**Returns:** `JSONContent` - TipTap document object

```typescript
const doc = parseHTML("<p>Hello World</p>");
```

#### `generateHTML(doc, extensions?)`

Generates an HTML string from TipTap JSON content.

**Parameters:**

- `doc: JSONContent` - TipTap document object
- `extensions?: Extensions` - Optional TipTap extensions

**Returns:** `string` - HTML string

```typescript
const html = generateHTML({ type: 'doc', content: [...] });
```

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

Use custom TipTap extensions for HTML/Markdown conversions:

```typescript
import { CustomExtension } from "./custom-extension";
import { parseHTML, generateHTML } from "docen";

const doc = parseHTML(html, [CustomExtension]);
const htmlContent = generateHTML(doc, [CustomExtension]);
```

### DOCX Template Patching

Replace `{{placeholders}}` in a DOCX template with TipTap-JSON content:

```typescript
import { patchDOCX, parseHTML, parseMarkdown } from "docen";

const result = await patchDOCX({
  template: templateBuffer,
  patches: {
    title: { content: parseHTML("<h1>Report</h1>") },
    body: { content: parseMarkdown("## Section\n\nHello **world**.") },
  },
  outputType: "nodebuffer",
});
```

Each patch's `content` is compiled to DOCX (styling derived from attrs) and the first section's children replace the placeholder. `keepOriginalStyles`, `recursive`, and `placeholderDelimiters` mirror the underlying `@office-open/docx` `patchDocument`.

## Format Conversion Matrix

| From \ To    | HTML     | Markdown | DOCX     |
| ------------ | -------- | -------- | -------- |
| **HTML**     | -        | via JSON | via JSON |
| **Markdown** | via JSON | -        | via JSON |
| **DOCX**     | via JSON | via JSON | -        |

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
- **Report Generation** - Generate DOCX reports from HTML/Markdown templates
- **Content Migration** - Migrate content between different formats
- **Collaborative Editing** - Use TipTap editor with format support

## Under the Hood

`docen` is a thin facade over [`@docen/docx`](../docx), the Tiptap DOCX editor package:

- **@docen/docx** - DOCX / HTML / Markdown converters built on the DocxManager architecture
- **@office-open/docx** - Native OOXML parse/generate (`parseDocument`, `generateDocument`, `patchDocument`)
- **@tiptap/html** / **@tiptap/markdown** - HTML and Markdown serialization (via @docen/docx)

## Comparison with Alternatives

| Feature     | docen | markdown-docx | mammoth | turndown |
| ----------- | ----- | ------------- | ------- | -------- |
| MD → DOCX   | ✅    | ✅            | ❌      | ❌       |
| DOCX → MD   | ✅    | ❌            | ❌      | ❌       |
| HTML ↔ MD   | ✅    | ❌            | ❌      | ✅       |
| DOCX ↔ HTML | ✅    | ❌            | ✅      | ❌       |
| TypeScript  | ✅    | ✅            | ✅      | ✅       |
| Unified API | ✅    | ❌            | ❌      | ❌       |
| Extensible  | ✅    | ❌            | ❌      | ✅       |

## License

MIT © [Demo Macro](https://www.demomacro.com/)
