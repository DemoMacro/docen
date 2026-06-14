You are a senior TypeScript developer.

## Project

**docen** is a monorepo for building online Office editors. Each editor package integrates its editing engine directly and uses `@docen/ui` (Fluent UI Web Components) for the UI layer. `@office-open/*` packages provide complete OOXML parse/generate APIs.

## Tech Stack

| Package | Engine                   | Parse/Generate    |
| ------- | ------------------------ | ----------------- |
| docx    | Tiptap (ProseMirror)     | @office-open/docx |
| pptx    | LeaferJS (Canvas)        | @office-open/pptx |
| xlsx    | RevoGrid (Data Grid)     | @office-open/xlsx |
| ui      | @fluentui/web-components | —                 |

## Monorepo Structure

```
packages/
  docx/      — @docen/docx (Tiptap DOCX editor + converters + UI)
  pptx/      — @docen/pptx (LeaferJS PPTX editor, future)
  xlsx/      — @docen/xlsx (RevoGrid XLSX editor, future)
  ui/        — @docen/ui (shared Fluent UI Web Components base layer)
```

- **@docen/ui** is a base UI component layer — toolbar, ribbon, buttons, menus. Editor packages consume it, not the other way around.
- Each editor package (docx, pptx, xlsx) directly integrates its engine and uses @docen/ui components for UI.

Legacy packages (kept for reference, not actively developed):

- `docen/`, `import-docx/`, `export-docx/`, `extensions/`, `deduplicate/`, `utils/`

## Build

- **Install**: `pnpm install`
- **Build**: `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one)
- **Build tool**: `vp pack` (vite-plus)
- **Lint**: `vp check`

## Data Model

**Dual-model architecture:**

| Role            | Format                              | Description                                             |
| --------------- | ----------------------------------- | ------------------------------------------------------- |
| **Runtime**     | Tiptap JSON (with DOCX-rich attrs)  | Editor directly operates on this via transactions       |
| **Persistence** | DocumentOptions (@office-open/docx) | Complete OOXML model for file I/O and format conversion |

Key principle: Tiptap JSON is the runtime model. DocumentOptions is the persistence/exchange model. Custom Tiptap extensions carry DOCX properties through `parseHTML`/`renderHTML` (node-level), `renderDocx`/`parseDocx`, and `attrs`.

## API Layering

**Standalone functions are core. Extension commands are thin wrappers.**

```typescript
// Format pipelines — runtime (Tiptap JSON) ↔ external formats
parseDOCX(buffer)                      → JSONContent               // DOCX → Tiptap JSON
generateDOCX<T>(json, options?)        → Promise<OutputByType[T]>  // prepare (default prepareImages) + compile + generateDocument; options.packer.type (default Buffer)
generateDOCXSync<T>(json, packer?)     → OutputByType[T]           // sync; compile + generateDocumentSync (no prepare — it's async)
generateDOCXStream(json, options?)     → Promise<ReadableStream>   // prepare + compile + generateDocumentStream
parseHTML(html)                        → JSONContent               // HTML → Tiptap JSON
generateHTML(json)                     → string                    // Tiptap JSON → HTML
parseMarkdown(md)                      → JSONContent               // Markdown → Tiptap JSON
generateMarkdown(json)                 → string                    // Tiptap JSON → Markdown

// Model bridge — runtime ↔ persistence via DocxManager (advanced; for layered control)
resolveDocument(docOpts)  → JSONContent       // DocumentOptions → Tiptap JSON
compileDocument(json)     → DocumentOptions   // Tiptap JSON → DocumentOptions
// Pre-compilation (http image → embedded data URL), in place. Required for http
// images — image renderDocx drops images without embedded data.
prepareDocument(json, steps?)           → Promise<void>            // default steps: [prepareImages()]
// parseDOCX / generateDOCX chain these internally over @office-open/docx's
// parseDocument / generateDocument:
//   buffer → parseDocument → resolve → JSON
//   JSON → prepare → compile → generateDocument → buffer
```

## Converter Pattern

Converters bridge the two models. `DocxManager` (in `converters/docx.ts`) walks the tree and assembles `DocumentOptions`. Extension modules contribute their DOCX expression per scope: single-node extensions export `renderDocx`/`parseDocx` (DocxManager dispatches per node); cross-node/container extensions (blockquote, lists, details, mention, task-item) export helper functions DocxManager orchestrates; simple-constant extensions (page-break, column-break) inline their payload. See CONTRIBUTING.md for the full scope table and helper inventory.

```typescript
// extensions/paragraph.ts — single-node extension exports renderDocx/parseDocx
export function renderDocx(node: JSONContent): ParagraphOptions
export function parseDocx(opts: ParagraphOptions): Record<string, unknown>

const manager = new DocxManager();
manager.compile(json)    → DocumentOptions  // Tiptap JSON → persistence
manager.resolve(docOpts) → JSONContent      // persistence → Tiptap JSON
```

Standalone functions (`resolveDocument`, `compileDocument`) use a default `DocxManager` instance internally.

## Extension Pattern

Custom Tiptap extensions extend the base `@tiptap/extension-*` to add DOCX-specific attributes:

- **attrs**: Define DOCX properties (shading, borders, indent, spacing, floating, crop, etc.) with `parseHTML` only
- **renderHTML** (node-level): Compute all CSS styles at once — solves the attribute-level style merge problem
- **renderDocx/parseDocx** (single-node extensions only): DOCX serialization (attrs ↔ DocumentOptions properties), exported as module-level functions and referenced from `extend({ renderDocx, parseDocx })`

Mark extensions (text-style, strike) keep attribute-level `renderHTML` since marks have a different rendering mechanism.

```typescript
export function renderDocx(node: JSONContent): ParagraphOptions { ... }
export function parseDocx(opts: ParagraphOptions): Record<string, unknown> { ... }

export const Paragraph = BaseParagraph.extend({
  addAttributes() { /* parseHTML only, no attribute-level renderHTML */ },
  renderHTML({ node, HTMLAttributes }) { /* node-level: all styles computed once */ },
  renderDocx,  // reference to exported function
  parseDocx,   // reference to exported function
});
```

Cross-node and simple-constant extensions use helper functions or inline payloads instead of `renderDocx`/`parseDocx` (see CONTRIBUTING.md).

This enables lossless round-trip through three serialization paths: HTML (ProseMirror DOMSerializer), DOCX (DocxManager), Markdown (MarkdownManager).

## Package Layout (packages/docx/src/)

```
src/
  index.ts        — Public API exports
  core.ts         — createDocxEditor(), docxExtensions config
  extensions/     — Custom Tiptap extensions (extends @tiptap/extension-*)
    utils.ts      — Shared renderHTML helpers (renderParagraphStyles, renderTableCellStyles)
  converters/     — DOCX, HTML, Markdown converters
    docx.ts       — DocxManager + parseDOCX/generateDOCX/resolveDocument/compileDocument
    html.ts       — parseHTML/generateHTML
    markdown.ts   — parseMarkdown/generateMarkdown
  types.ts        — Public type definitions
```

## Naming Conventions

- **Functions**: `parse*`/`generate*` (external-format I/O — JSONContent ↔ DOCX/HTML/Markdown), `resolve*` (DocOpts → JSON), `compile*` (JSON → DocOpts)
- **Files**: kebab-case
- **Interfaces**: PascalCase without `I` prefix, `Options` suffix, `readonly` properties
- **Constants**: `as const` objects, SCREAMING_SNAKE_CASE keys
- **Loops**: `for...of` default, `.map()` only when returning new array

## Performance

- Large documents: custom pagination using Pretext (text measurement) + Paged.js (print rendering)
- Disable spellcheck for very large documents
- Pagination is a second-phase feature; current focus is core converters

## Behavioral Guidelines

- State assumptions explicitly. If uncertain, ask before implementing.
- No features beyond what was asked. No speculative abstractions.
- Touch only what you must. Match existing style.
- Transform tasks into verifiable goals. Loop until verified.
