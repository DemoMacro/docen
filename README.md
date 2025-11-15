# TipTap Extension

![GitHub](https://img.shields.io/github/license/DemoMacro/tiptap-extension)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> A collection of powerful extensions for TipTap/ProseMirror editor, including DOCX export functionality and more.

## Packages

- **[tiptap-extension-export-docx](./packages/export-docx)** - Export TipTap/ProseMirror content to Microsoft Word DOCX format

## Quick Start

### DOCX Export

```bash
# Install with npm
$ npm install tiptap-extension-export-docx

# Install with yarn
$ yarn add tiptap-extension-export-docx

# Install with pnpm
$ pnpm add tiptap-extension-export-docx
```

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

## Development

### Prerequisites

- **Node.js** 18.x or higher
- **pnpm** 9.x or higher (recommended package manager)
- **Git** for version control

### Getting Started

1. **Clone the repository**:

   ```bash
   git clone https://github.com/DemoMacro/tiptap-extension.git
   cd tiptap-extension
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
   import { generateDOCX } from 'tiptap-extension-export-docx';
   ```

### Development Commands

```bash
pnpm dev            # Development mode with watch
pnpm build          # Build all packages
pnpm lint           # Run code formatting and linting
```

## Repository Structure

```
tiptap-extension/
├── packages/
│   └── export-docx/              # DOCX export extension
│       ├── src/
│       │   ├── docx.ts           # Main export functionality
│       │   ├── converters/       # Content converters
│       │   ├── types.ts          # Type definitions
│       │   ├── option.ts         # Configuration options
│       │   └── utils.ts          # Utility functions
│       ├── package.json          # Package manifest
│       ├── build.config.ts       # Build configuration
│       └── README.md             # Package documentation
├── playground/                    # Testing and examples
│   ├── html/                     # HTML test files
│   ├── extensions.ts             # TipTap extensions
│   └── index.ts                  # Test runner
├── package.json                  # Monorepo configuration
├── pnpm-workspace.yaml           # pnpm workspace configuration
└── README.md                    # This file
```

## Architecture & Design

### 🏗️ **Extension Architecture**

- **Converters** (`src/converters/`): Content type to DOCX conversion
- **Types** (`src/types.ts`): TypeScript type definitions for TipTap content
- **Options** (`src/option.ts`): Configuration system for export settings
- **Utils** (`src/utils.ts`): Shared utilities for image processing and helpers

### 🔧 **Technology Stack**

- **[docx](https://docx.js.org/)**: Microsoft Word document generation
- **[ofetch](https://github.com/unjs/ofetch)**: Modern fetch API with better error handling
- **[image-meta](https://github.com/unjs/image-meta)**: Image metadata extraction
- **[unbuild](https://github.com/unjs/unbuild)**: Modern bundler and build tool

### 📝 **Configuration System**

The export system supports extensive configuration:

- **Document Properties**: Title, author, description, metadata
- **Image Options**: Default sizing, positioning, text wrapping
- **Table Options**: Layout, borders, cell formatting
- **Output Types**: Multiple formats (Buffer, ArrayBuffer, Base64, etc.)

## Contributing

We welcome contributions! Here's how to get started:

### Quick Setup

1. **Fork the repository** on GitHub
2. **Clone your fork**:

   ```bash
   git clone https://github.com/YOUR_USERNAME/tiptap-extension.git
   cd tiptap-extension
   ```

3. **Add upstream remote**:

   ```bash
   git remote add upstream https://github.com/DemoMacro/tiptap-extension.git
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
   import { generateDOCX } from 'tiptap-extension-export-docx';
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

- 📫 [Report Issues](https://github.com/DemoMacro/tiptap-extension/issues)
- 📚 [Package Documentation](./packages/export-docx/README.md)

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

---

Built with ❤️ by [Demo Macro](https://imst.xyz/)
