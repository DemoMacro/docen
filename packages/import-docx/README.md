# @docen/import-docx

![npm version](https://img.shields.io/npm/v/@docen/import-docx)
![npm downloads](https://img.shields.io/npm/dw/@docen/import-docx)
![npm license](https://img.shields.io/npm/l/@docen/import-docx)

> Import Microsoft Word DOCX files to TipTap/ProseMirror content.

## Features

- 📝 **Rich Text Parsing** - Accurate parsing of headings, paragraphs, and blockquotes with formatting
- 🖼️ **Image Extraction** - Automatic image extraction and base64 conversion
- 📊 **Table Support** - Complete table structure with colspan/rowspan detection algorithm
- ✅ **Lists & Tasks** - Bullet lists, numbered lists with start number extraction, and task lists with checkbox detection
- 🎨 **Text Formatting** - Bold, italic, underline, strikethrough, subscript, superscript, and highlights
- 🎯 **Text Styles** - Comprehensive style support including colors, backgrounds, fonts, sizes, and line heights
- 🔗 **Links** - Hyperlink extraction with href preservation
- 💻 **Code Blocks** - Code block detection with language attribute extraction
- 🧠 **Smart Parsing** - DOCX XML parsing with proper element grouping and structure reconstruction
- ⚡ **Fast Processing** - Uses fflate for ultra-fast ZIP decompression

## Installation

```bash
# Install with npm
$ npm install @docen/import-docx

# Install with yarn
$ yarn add @docen/import-docx

# Install with pnpm
$ pnpm add @docen/import-docx
```

## Quick Start

```typescript
import { parseDOCX } from "@docen/import-docx";
import { readFileSync } from "node:fs";

// Read DOCX file
const buffer = readFileSync("document.docx");

// Parse DOCX to TipTap JSON
const content = await parseDOCX(buffer);

// Use in TipTap editor
editor.commands.setContent(content);
```

## API Reference

### `parseDOCX(input, options?)`

Parses a DOCX file and converts it to TipTap/ProseMirror JSON content.

**Parameters:**

- `input: Buffer | ArrayBuffer | Uint8Array` - DOCX file data
- `options?: DocxImportOptions` - Optional import configuration

**Returns:** `Promise<JSONContent>` - TipTap/ProseMirror document content with images embedded

**Options:**

```typescript
interface DocxImportOptions {
  /** Custom image converter (default: embed as base64) */
  convertImage?: (image: DocxImageInfo) => Promise<DocxImageResult>;

  /** Whether to ignore empty paragraphs (default: false) */
  ignoreEmptyParagraphs?: boolean;
}
```

**Default Image Converter:**

The package exports `defaultImageConverter` which embeds images as base64 data URLs:

```typescript
import { defaultImageConverter } from "@docen/import-docx";

// Use in custom converter
await parseDOCX(buffer, {
  convertImage: async (image) => {
    if (shouldUploadToCDN) {
      return uploadToCDN(image.data);
    }
    return defaultImageConverter(image);
  },
});
```

## Supported Content Types

### Text Formatting

- **Bold**, _Italic_, <u>Underline</u>, ~~Strikethrough~~
- ^Superscript^ and ~Subscript~
- Text highlights
- Text colors and background colors
- Font families and sizes
- Line heights

### Block Elements

- **Headings** (H1-H6) with proper level detection
- **Paragraphs** with text alignment (left, right, center, justify)
- **Blockquotes** (Detected by indentation + left border formatting)
- **Horizontal Rules** (Detected as page breaks in DOCX)
- **Code Blocks** with language attribute support

### Lists

- **Bullet Lists** with proper nesting and structure
- **Numbered Lists** with custom start number extraction
- **Task Lists** with checked/unchecked state detection (☐/☑ symbols)

### Tables

- Complete table structure parsing
- **Table Cells** with colspan detection using grid-based algorithm
- **Table Cells** with rowspan detection using vMerge tracking
- Cell alignment and formatting preservation
- Merged cell handling (both horizontal and vertical)

### Media & Embeds

- **Images** with automatic base64 conversion
- **Links** (hyperlinks) with href extraction

## Parsing Algorithm

### Document Structure

The parser follows a structured workflow:

1. **Extract Relationships** - Parse `_rels/document.xml.rels` for hyperlinks and images
2. **Parse Numbering** - Extract list definitions from `numbering.xml` (abstractNum → numFmt)
3. **Process Document Body** - Iterate through document.xml elements:
   - Detect content types (tables, lists, paragraphs, code blocks, etc.)
   - Group consecutive elements into proper containers
   - Convert XML nodes to TipTap JSON nodes

### Table Processing

Tables use specialized algorithms:

- **Colspan Detection** - Grid-based algorithm tracks cell positions and detects horizontal merges
- **Rowspan Detection** - Vertical merge (vMerge) tracking across rows with proper cell skipping
- **Cell Content** - Recursive parsing of nested paragraphs and formatting
- **Hyperlink Support** - Proper handling of links within table cells

### List Processing

Lists utilize the DOCX numbering system:

- **Numbering ID Mapping** - Maps abstractNum to formatting (bullet vs decimal)
- **Start Value Extraction** - Extracts and preserves start numbers for ordered lists
- **Nesting Preservation** - Maintains proper list hierarchy
- **Consecutive Grouping** - Groups consecutive list items into list containers

## Examples

### Basic Usage

```typescript
import { parseDOCX } from "@docen/import-docx";

const buffer = readFileSync("example.docx");
const { content } = await parseDOCX(buffer);

console.log(JSON.stringify(content, null, 2));
```

### Use with TipTap Editor

```typescript
import { Editor } from "@tiptap/core";
import { parseDOCX } from "@docen/import-docx";

const editor = new Editor({
  extensions: [...],
  content: "",
});

// Import DOCX file
async function importDocx(file: File) {
  const buffer = await file.arrayBuffer();
  const content = await parseDOCX(buffer);
  editor.commands.setContent(content);
}
```

## Known Limitations

### Blockquote Detection

DOCX does not have a semantic blockquote structure. Blockquotes are detected by:

- Left indentation ≥ 720 twips (0.5 inch)
- Presence of left border (single line)

This detection method may produce false positives for documents with custom indentation similar to blockquotes.

### Code Marks

The `code` mark is NOT automatically detected from monospace fonts (Consolas, Courier New, etc.). This is intentional to avoid false positives. Code marks should be explicitly added in the source document or through editor UI.

### Color Format

All colors are imported as hex values (e.g., "#FF0000", "#008000"). Color names from the original document are not preserved.

### Image Limitations

- Only embedded images are supported (external image links are not fetched)
- Image width/height metadata is preserved but visual sizing may vary
- Title, alt text, and other image attributes have limited DOCX support

### Table Cell Types

DOCX format does not distinguish between header and body cells at a semantic level. All cells are imported as `tableCell` type for consistency. This is a DOCX format limitation.

## Contributing

Contributions are welcome! Please read our [Contributor Covenant](https://www.contributor-covenant.org/version/2/1/code_of_conduct/) and submit pull requests to the [main repository](https://github.com/DemoMacro/docen).

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
