> **⚠️ Warning:** This project is not yet stable and may undergo significant changes before reaching version 1.0.0. We strongly advise against using it in production environments.

# Docen

![GitHub](https://img.shields.io/github/license/DemoMacro/docen)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> Universal document format converter and DOCX editor built on TipTap/ProseMirror, with comprehensive TypeScript support. Convert between Markdown, HTML, and DOCX through a unified Tiptap JSON model.

## Packages

| Package                                  | Version                                          | Description                                                         |
| ---------------------------------------- | ------------------------------------------------ | ------------------------------------------------------------------- |
| [docen](./packages/docen/README.md)      | ![npm](https://img.shields.io/npm/v/docen)       | Universal converter — unified JSON API for Markdown, HTML, and DOCX |
| [@docen/docx](./packages/docx/README.md) | ![npm](https://img.shields.io/npm/v/@docen/docx) | Tiptap DOCX editor + converters, powered by @office-open/docx       |

## Quick Start

### Universal Converter (`docen`)

For seamless conversion between Markdown, HTML, and DOCX through a single unified API:

```bash
# Install with pnpm
$ pnpm add docen
```

```typescript
import { parseHTML, generateDOCX, parseMarkdown, generateHTML } from "docen";

// HTML → DOCX
const doc = parseHTML("<h1>Title</h1><p>Hello World</p>");
const docx = await generateDOCX(doc);

// Markdown → HTML
const doc2 = parseMarkdown("# Title\n\nHello World");
const html = generateHTML(doc2);
```

### DOCX Editor (`@docen/docx`)

A full-featured WYSIWYG DOCX editor with near-lossless round-trip conversion:

```bash
$ pnpm add @docen/docx
```

```typescript
import { createDocxEditor, parseDOCX, generateDOCX } from "@docen/docx";

const editor = createDocxEditor({ element: document.querySelector("#editor") });
editor.commands.setContent(parseDOCX(buffer));
const output = await generateDOCX(editor.getJSON());
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

3. **Build all packages**:

   ```bash
   pnpm build
   ```

### Development Commands

```bash
pnpm build                       # Build all packages
cd packages/<pkg> && pnpm build  # Build one package
vp check                         # Lint & format
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

5. **Build**:

   ```bash
   pnpm build
   ```

### Development Workflow

1. **Code**: Follow our project standards (see [CONTRIBUTING.md](./CONTRIBUTING.md))
2. **Test**: `pnpm build && <verify your changes>`
3. **Commit**: Use conventional commits (`feat:`, `fix:`, `docs:`, `style:`, `refactor:`, `perf:`, `test:`, `build:`, `ci:`, `chore:`, `revert:`)
4. **Push**: Push to your fork
5. **Submit**: Create a Pull Request to upstream repository

## Support & Community

- 📫 [Report Issues](https://github.com/DemoMacro/docen/issues)
- 📚 [docen Documentation](./packages/docen/README.md)
- 📚 [@docen/docx Documentation](./packages/docx/README.md)

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

---

Built with ❤️ by [Demo Macro](https://www.demomacro.com/)
