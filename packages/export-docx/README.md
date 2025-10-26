# TipTap Extension - DOCX Export

![npm version](https://img.shields.io/npm/v/tiptap-extension-export-docx)
![npm downloads](https://img.shields.io/npm/dw/tiptap-extension-export-docx)
![npm license](https://img.shields.io/npm/l/tiptap-extension-export-docx)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> A powerful TipTap/ProseMirror extension that converts editor content to Microsoft Word DOCX format.

## Features

- 📝 **Rich Text Support** - Headings, paragraphs, lists, blockquotes
- 🖼️ **Image Handling** - Automatic sizing and metadata extraction
- 📊 **Table Support** - Complete table structure with headers
- 📋 **Lists & Tasks** - Bullet lists, numbered lists, and task lists
- 🎨 **Formatting** - Bold, italic, underline, strikethrough, colors
- 🔗 **Links** - Hyperlink support
- 💻 **Code Blocks** - Syntax highlighted code blocks
- ⚙️ **Configurable** - Customizable export options

## Installation

```bash
# Install with npm
$ npm install tiptap-extension-export-docx

# Install with yarn
$ yarn add tiptap-extension-export-docx

# Install with pnpm
$ pnpm add tiptap-extension-export-docx
```

## Usage

```typescript
import { generateDOCX } from "tiptap-extension-export-docx";
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
- `options: DocxOptions` - Export configuration options

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

### `DocxOptions`

Configuration options for DOCX generation:

📖 **See complete interface definition**: [option.ts](./src/option.ts)

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
