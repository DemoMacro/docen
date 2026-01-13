# @docen/export-docx

![npm version](https://img.shields.io/npm/v/@docen/export-docx)
![npm downloads](https://img.shields.io/npm/dw/@docen/export-docx)
![npm license](https://img.shields.io/npm/l/@docen/export-docx)

> Export TipTap/ProseMirror editor content to Microsoft Word DOCX format.

## Features

- 📝 **Rich Text Support** - Full support for headings, paragraphs, and blockquotes with proper formatting
- 🖼️ **Image Handling** - Automatic image sizing, positioning, and metadata extraction
- 📊 **Table Support** - Complete table structure with headers, cells, colspan, and rowspan
- ✅ **Lists & Tasks** - Bullet lists, numbered lists with custom start numbers, and task lists with checkboxes
- 🎨 **Text Formatting** - Bold, italic, underline, strikethrough, subscript, and superscript
- 🎯 **Text Styles** - Comprehensive style support including colors, backgrounds, fonts, sizes, and line heights
- 🔗 **Links** - Hyperlink support with href preservation
- 💻 **Code Blocks** - Syntax highlighted code blocks with language attribute support
- 📁 **Collapsible Content** - Details/summary sections for expandable content
- 😀 **Emoji Support** - Native emoji rendering in documents
- 🧮 **Mathematical Content** - LaTeX-style formula support
- ⚙️ **Configurable Options** - Customizable export options for documents, tables, styles, and horizontal rules

## Installation

```bash
# Install with npm
$ npm install @docen/export-docx

# Install with yarn
$ yarn add @docen/export-docx

# Install with pnpm
$ pnpm add @docen/export-docx
```

## Quick Start

```typescript
import { generateDOCX } from "@docen/export-docx";
import { writeFileSync } from "node:fs";

// Your TipTap/ProseMirror editor content
const content = {
  type: "doc",
  content: [
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          marks: [{ type: "bold" }, { type: "italic" }],
          text: "Hello, world!",
        },
      ],
    },
  ],
};

// Convert to DOCX and save to file
const docx = await generateDOCX(content, { outputType: "nodebuffer" });
writeFileSync("document.docx", docx);
```

## API Reference

### `generateDOCX(content, options)`

Converts TipTap/ProseMirror content to DOCX format.

**Parameters:**

- `content: JSONContent` - TipTap/ProseMirror editor content
- `options: DocxExportOptions` - Export configuration options

**Returns:** `Promise<OutputByType[T]>` - DOCX file data with type matching the specified outputType

**Available Output Types:**

- `"base64"` - Base64 encoded string
- `"string"` - Text string
- `"text"` - Plain text
- `"binarystring"` - Binary string
- `"array"` - Array of numbers
- `"uint8array"` - Uint8Array
- `"arraybuffer"` - ArrayBuffer
- `"blob"` - Blob object
- `"nodebuffer"` - Node.js Buffer

**Configuration Options:**

- `title` - Document title
- `creator` - Document author
- `description` - Document description
- `outputType` - Output format (required)
- `table` - Table styling defaults (alignment, spacing, borders)
- `image` - Image handling options
- `styles` - Document default styles (font, line height, spacing)
- `horizontalRule` - Horizontal rule style

## Supported Content Types

### Text Formatting

- **Bold**, _Italic_, <u>Underline</u>, ~~Strikethrough~~
- ^Superscript^ and ~Subscript~
- Text colors and background colors
- Font families and sizes
- Line heights

### Block Elements

- **Headings** (H1-H6) with level attribute
- **Paragraphs** with text alignment (left, right, center, justify)
- **Blockquotes** (Note: Exported as indented paragraphs with left border due to DOCX format)
- **Horizontal Rules** (Exported as page breaks by default)
- **Code Blocks** with language support

### Lists

- **Bullet Lists** - Standard unordered lists
- **Numbered Lists** - Ordered lists with custom start number
- **Task Lists** - Checkbox lists with checked/unchecked states

### Tables

- Complete table structure with rows and cells
- **Table Headers** with colspan/rowspan support
- **Table Cells** with colspan/rowspan support
- Cell alignment and formatting options

### Media & Embeds

- **Images** with automatic sizing and positioning
- **Links** (hyperlinks) with href attribute
- **Emoji** rendering
- **Mathematics** formulas (LaTeX-style)
- **Details/Summary** collapsible sections

## Examples

### Document with Tables and Colspan/Rowspan

```typescript
const content = {
  type: "doc",
  content: [
    {
      type: "table",
      content: [
        {
          type: "tableRow",
          content: [
            {
              type: "tableHeader",
              attrs: { colspan: 2, rowspan: 1 },
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "text", text: "Spanning Header" }],
                },
              ],
            },
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "text", text: "Regular Cell" }],
                },
              ],
            },
          ],
        },
      ],
    },
  ],
};
```

### Document with Text Styles

```typescript
const content = {
  type: "doc",
  content: [
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          marks: [
            {
              type: "textStyle",
              attrs: {
                color: "#FF0000",
                fontSize: "18px",
                fontFamily: "Arial",
                backgroundColor: "#FFFF00",
              },
            },
          ],
          text: "Red, 18px, Arial text on yellow background",
        },
      ],
    },
  ],
};
```

### Document with Lists

```typescript
const content = {
  type: "doc",
  content: [
    {
      type: "bulletList",
      content: [
        {
          type: "listItem",
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "First item" }],
            },
          ],
        },
        {
          type: "listItem",
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "Second item" }],
            },
          ],
        },
      ],
    },
  ],
};
```

## Known Limitations

### Blockquote Structure

DOCX does not have a semantic blockquote structure. Blockquotes are exported as:

- Indented paragraphs (720 twips / 0.5 inch left indentation)
- Left border (single line)

This is a DOCX format limitation, not a bug.

### Code Marks

The `code` mark is exported as monospace font (Consolas). When re-importing, it will be recognized as `textStyle` with `fontFamily: "Consolas"`, not as a `code` mark.

This is intentional - we do not detect code marks from fonts during import to avoid false positives.

### Color Name Conversion

Color names (like `"red"`, `"green"`, `"blue"`) are automatically converted to hex values (`"#FF0000"`, `"#008000"`, `"#0000FF"`) for DOCX compatibility.

## Contributing

Contributions are welcome! Please read our [Contributor Covenant](https://www.contributor-covenant.org/version/2/1/code_of_conduct/) and submit pull requests to the [main repository](https://github.com/DemoMacro/docen).

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
