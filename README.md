# Docen

![GitHub](https://img.shields.io/github/license/DemoMacro/docen)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> A modern rich-text editor based on TipTap with comprehensive extensions for document processing.

## Packages

- **[@docen/extensions](./packages/extensions)** - TipTap extensions with comprehensive TypeScript types
- **[@docen/export-docx](./packages/export-docx)** - Export TipTap/ProseMirror content to Microsoft Word DOCX format
- **[@docen/import-docx](./packages/import-docx)** - Import Microsoft Word DOCX files to TipTap/ProseMirror content

## Quick Start

### DOCX Export

```bash
# Install with npm
$ npm install @docen/export-docx

# Install with yarn
$ yarn add @docen/export-docx

# Install with pnpm
$ pnpm add @docen/export-docx
```

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

### DOCX Import

```bash
# Install with npm
$ npm install @docen/import-docx

# Install with yarn
$ yarn add @docen/import-docx

# Install with pnpm
$ pnpm add @docen/import-docx
```

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

## Development

### Prerequisites

- **Node.js** 18.x or higher
- **pnpm** 9.x or higher (recommended package manager)
- **Git** for version control

### Getting Started

1. **Clone the repository**:

   ```bash
   git clone https://github.com/DemoMacro/docen.git
   cd docen
   ```

2. **Install dependencies**:

   ```bash
   pnpm install
   ```

3. **Development mode**:

   ```bash
   pnpm dev
   ```

4. **Build all packages**:

   ```bash
   pnpm build
   ```

5. **Test locally**:

   ```bash
   # Link the package globally for testing
   cd packages/export-docx
   pnpm link --global

   # Test in your project
   import { generateDOCX } from '@docen/export-docx';
   ```

### Development Commands

```bash
pnpm dev            # Development mode with watch
pnpm build          # Build all packages
pnpm lint           # Run code formatting and linting
```

## Contributing

We welcome contributions! Here's how to get started:

### Quick Setup

1. **Fork the repository** on GitHub
2. **Clone your fork**:

   ```bash
   git clone https://github.com/YOUR_USERNAME/docen.git
   cd docen
   ```

3. **Add upstream remote**:

   ```bash
   git remote add upstream https://github.com/DemoMacro/docen.git
   ```

4. **Install dependencies**:

   ```bash
   pnpm install
   ```

5. **Development mode**:

   ```bash
   pnpm dev
   ```

6. **Test locally**:

   ```bash
   # Link the package globally for testing
   cd packages/export-docx
   pnpm link --global

   # Test your changes
   import { generateDOCX } from '@docen/export-docx';
   ```

### Development Workflow

1. **Code**: Follow our project standards
2. **Test**: `pnpm build && <test your extension>`
3. **Commit**: Use conventional commits (`feat:`, `fix:`, etc.)
4. **Push**: Push to your fork
5. **Submit**: Create a Pull Request to upstream repository

## Project Philosophy

This project follows core principles:

1. **TipTap Focus**: Built specifically for TipTap/ProseMirror ecosystem
2. **Type Safety**: Full TypeScript support with comprehensive types
3. **Modular Design**: Each converter handles specific content types
4. **Extensible**: Easy to add new content type converters
5. **Performance**: Optimized for large documents and batch processing
6. **User Experience**: Simple API with powerful configuration options

## Support & Community

- 📫 [Report Issues](https://github.com/DemoMacro/docen/issues)
- 📚 [Export Documentation](./packages/export-docx/README.md)
- 📚 [Import Documentation](./packages/import-docx/README.md)

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

---

Built with ❤️ by [Demo Macro](https://imst.xyz/)
