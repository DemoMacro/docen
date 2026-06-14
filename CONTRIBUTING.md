# Contributing to docen

Thank you for your interest in contributing! This document describes the coding standards and conventions.

## Development Setup

```bash
pnpm install            # Install dependencies
pnpm build              # Build all packages
cd packages/<pkg> && pnpm build   # Build one package
vp check                # Lint
```

## Project Structure

```
packages/
  docx/     — @docen/docx (Tiptap DOCX editor + converters + UI)
  pptx/     — @docen/pptx (LeaferJS PPTX editor, future)
  xlsx/     — @docen/xlsx (RevoGrid XLSX editor, future)
  ui/       — @docen/ui (shared Fluent UI Web Components base layer)
```

- **@docen/ui** is a base UI component layer (toolbar, ribbon, buttons, menus). Editor packages consume it.
- Each editor package directly integrates its engine and uses @docen/ui components for UI.

Each editor package follows the same layout:

```
src/
  index.ts        — Public API
  core.ts         — Editor factory and extension configuration
  extensions/     — Custom Tiptap extensions (extends @tiptap/extension-*)
    utils.ts      — Shared renderHTML helpers
  converters/     — DOCX, HTML, Markdown converters
    docx.ts       — DocxManager + standalone functions
    html.ts       — parseHTML/generateHTML
    markdown.ts   — parseMarkdown/generateMarkdown
  types.ts        — Public type definitions
```

## Naming Conventions

### Functions

Use **camelCase**. Follow the appropriate prefix convention:

| Prefix      | Purpose                               | Example                                                  |
| ----------- | ------------------------------------- | -------------------------------------------------------- |
| `parse*`    | External format → runtime/persistence | `parseDOCX()`, `parseHTML()`, `parseMarkdown()`          |
| `generate*` | runtime/persistence → External format | `generateDOCX()`, `generateHTML()`, `generateMarkdown()` |
| `resolve*`  | DocumentOptions → Tiptap JSON         | `resolveDocument()`, `resolveParagraph()`                |
| `compile*`  | Tiptap JSON → DocumentOptions         | `compileDocument()`, `compileParagraph()`                |
| `create*`   | Factory functions                     | `createDocxEditor()`                                     |

### Files and Directories

Use **kebab-case** for all file and directory names.

```
extensions/paragraph.ts    — custom extension with renderDocx/parseDocx
converters/docx.ts         — DocxManager + standalone functions
converters/html.ts         — HTML conversion
```

### Interfaces

**PascalCase** without `I` prefix. Configuration interfaces use `Options` suffix. All properties `readonly`.

```typescript
export interface ParagraphNode {
  type: "paragraph";
  attrs?: {
    readonly textAlign?: "left" | "center" | "right" | "justify";
    readonly shading?: Shading;
  };
  content?: Array<TextNode | HardBreakNode>;
}
```

### Constants (Enumerated Types)

Use `as const` objects (not TypeScript `enum`). Keys use **SCREAMING_SNAKE_CASE**. Values use **lowercase**.

```typescript
export const AlignmentType = {
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
  JUSTIFY: "justify",
} as const;
```

## Converter Design Pattern

Converters bridge Tiptap JSON (runtime) and DocumentOptions (persistence). `DocxManager` (in `converters/docx.ts`) walks the tree and assembles `DocumentOptions`. Extension modules contribute their DOCX expression in one of three ways, depending on scope:

| Scope                      | Extensions                                                         | DOCX contribution                                                                                        |
| -------------------------- | ------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------- |
| **Single-node**            | paragraph, heading, image, code-block, table\*, text-style, strike | export `renderDocx(node)`/`parseDocx(opts)` — DocxManager dispatches per node                            |
| **Cross-node / container** | blockquote, ordered-list, task-item, mention, details              | export helpers (`buildOrderedLevels`, `createMention`, …) — DocxManager orchestrates multi-node assembly |
| **Simple constant**        | page-break, column-break                                           | payload inlined in `DocxManager` (single line, no variance)                                              |

```typescript
// DocxManager — central dispatcher
const manager = new DocxManager();
manager.compile(json)    → DocumentOptions   // Tiptap JSON → persistence
manager.resolve(docOpts) → JSONContent       // persistence → Tiptap JSON

// Single-node extensions export renderDocx/parseDocx (DocxManager dispatches)
// extensions/paragraph.ts
export function renderDocx(node: JSONContent): ParagraphOptions { ... }
export function parseDocx(opts: ParagraphOptions): Record<string, unknown> { ... }
```

Standalone functions use a default `DocxManager` instance internally:

```typescript
// resolveDocument = defaultManager.resolve(docOpts)
// compileDocument = defaultManager.compile(json)
export function resolveDocument(docOpts: DocumentOptions): JSONContent;
export function compileDocument(json: JSONContent): DocumentOptions;
```

`parseDOCX`/`generateDOCX` are high-level JSON APIs — like `parseHTML`/`parseMarkdown`, they operate directly on Tiptap JSON (not `DocumentOptions`). `generateDOCX` runs `prepareDocument → compileDocument → generateDocument` internally:

```typescript
// parseDOCX = parseDocument (office-open) → resolveDocument → Tiptap JSON
export function parseDOCX(buffer): JSONContent;

// generateDOCX = prepareDocument → compileDocument → generateDocument (office-open)
export function generateDOCX(json): Promise<Buffer>;
```

`resolveDocument`/`compileDocument` (the `DocumentOptions` ↔ Tiptap JSON model bridge) are advanced internals — rarely needed directly.

## Extension Design Pattern

Custom Tiptap extensions extend the base `@tiptap/extension-*` to carry DOCX-specific properties. Each extension:

1. **Adds attrs** with `parseHTML` only (attribute-level rendering is not used for nodes)
2. **Defines node-level `renderHTML`** to compute all CSS styles at once (avoids style merge conflicts)
3. **Exports `renderDocx`/`parseDocx`** for DOCX serialization (used by DocxManager) — this applies to **single-node** extensions only (see the three-scope table above)

Mark extensions (text-style, strike) keep attribute-level `renderHTML` since marks have a different rendering mechanism.

```typescript
// Exported for DocxManager
export function renderDocx(node: JSONContent): ParagraphOptions { ... }
export function parseDocx(opts: ParagraphOptions): Record<string, unknown> { ... }

export const Paragraph = BaseParagraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      indentLeft: {
        default: null,
        parseHTML: (element) => element.style.marginLeft || null,
        // No renderHTML — node-level renderHTML handles all styles
      },
    };
  },
  renderHTML({ node, HTMLAttributes }) {
    const styles = renderParagraphStyles(node.attrs);
    const attrs = { ...HTMLAttributes };
    if (styles.length > 0) attrs.style = styles.join(";");
    return ["p", attrs, 0] as const;
  },
  renderDocx,
  parseDocx,
});
```

Cross-node extensions (blockquote, ordered-list, task-item, mention, details) carry no `renderDocx`/`parseDocx` — their DOCX expression spans multiple nodes, so they export helper functions (`applyBlockquoteStyle`, `buildOrderedLevels`, `createTaskCheckbox`, `createMention`, details constants) that `DocxManager` calls during tree assembly. Simple-constant extensions (page-break, column-break) inline their one-line DOCX payload directly in `DocxManager`.

## API Layering

**Standalone functions are core. Extension commands are thin wrappers.**

- Standalone functions (`parseDOCX`, `generateDOCX`, etc.) work without an editor instance — for headless/server/batch use
- Tiptap extension commands (`editor.commands.importDocx()`, etc.) are convenience wrappers that call standalone functions internally
- @docen/ui components call either layer as appropriate

## Loop Patterns

| Scenario                            | Use                 | Reason                            |
| ----------------------------------- | ------------------- | --------------------------------- |
| Transform into new array            | `.map()`            | Expresses "transform" intent      |
| Filter elements                     | `.filter()`         | Expresses "filter" intent         |
| Side-effect iteration, async, break | `for...of`          | Full control, supports early exit |
| Performance-sensitive hot paths     | `for...of` or `for` | ~3x faster than `.forEach()`      |

**Avoid `.forEach()`** — `for...of` is strictly superior.

## Pull Request Process

1. `vp check` passes with no errors
2. `pnpm build` succeeds for the changed package
3. Follow naming conventions described above
4. Keep changes minimal and focused — match existing style
