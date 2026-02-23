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
const doc = await parseDOCX(buffer);

// Generate DOCX from TipTap JSON
const docxBuffer = await generateDOCX(doc, { outputType: "nodebuffer" });
```

### Cross-Format Conversion

Convert between any formats by using TipTap JSON as the intermediate format:

```typescript
import { parseHTML, generateDOCX } from "docen";

// HTML → DOCX
const html = "<h1>Title</h1><p>Content...</p>";
const doc = parseHTML(html);
const docx = await generateDOCX(doc, { outputType: "blob" });

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
- `extensions?: Extensions` - Optional TipTap extensions (defaults to @docen/extensions)
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

#### `parseDOCX(input, options?)`

Parses a DOCX file into TipTap JSON content.

**Parameters:**

- `input: Buffer | ArrayBuffer | Uint8Array | string` - DOCX file data or path
- `options?: DocxImportOptions` - Optional import configuration

**Returns:** `Promise<JSONContent>` - TipTap document object

```typescript
import { readFileSync } from "node:fs";
const buffer = readFileSync("document.docx");
const doc = await parseDOCX(buffer);
```

#### `generateDOCX(docJson, options)`

Generates a DOCX file from TipTap JSON content.

**Parameters:**

- `docJson: JSONContent` - TipTap document object
- `options: DocxExportOptions` - Export options including outputType

**Returns:** `Promise<ArrayBuffer | Uint8Array | Blob | Buffer>` - DOCX data in specified format

```typescript
const buffer = await generateDOCX(doc, {
  outputType: "nodebuffer",
  title: "My Document",
});
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

### DOCX Export Options

Customize DOCX generation with advanced options:

```typescript
const docx = await generateDOCX(doc, {
  outputType: "blob",
  title: "Document Title",
  creator: "Author Name",
  styles: {
    default: {
      document: {
        paragraph: { spacing: { line: 480 } },
        run: { size: 28 },
      },
    },
  },
  image: {
    maxWidth: 600,
    handler: customImageHandler,
  },
  table: {
    run: {
      width: { size: 100, type: "pct" },
      alignment: "center",
    },
  },
});
```

### DOCX Import Options

Control DOCX parsing behavior:

```typescript
const doc = await parseDOCX(buffer, {
  ignoreEmptyParagraphs: true,
  enableImageCrop: true,
  canvasImport: () => import("@napi-rs/canvas"),
  image: {
    handler: async (image) => {
      // Custom image handling (upload to CDN, etc.)
      return await uploadToCloud(image.data);
    },
  },
});
```

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

`docen` leverages the powerful TipTap/ProseMirror ecosystem:

- **@tiptap/html** - HTML parsing and generation
- **@tiptap/markdown** - Markdown parsing and generation with MarkdownManager
- **@docen/import-docx** - DOCX parsing with advanced XML processing
- **@docen/export-docx** - DOCX generation with docx library
- **@docen/extensions** - Comprehensive TipTap extension collection

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

MIT © [Demo Macro](https://imst.xyz/)
