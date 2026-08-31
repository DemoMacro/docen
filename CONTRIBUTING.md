# Contributing to docen

Thanks for contributing! This guide covers the **workflow** for contributing and the **coding standards** that keep docen consistent. For architectural context (data models, API layering, design decisions), see [CLAUDE.md](./CLAUDE.md).

## Development Setup

```bash
pnpm install                                # install dependencies
pnpm build                                  # build all packages
cd packages/<pkg> && pnpm build             # build one package
cd packages/<pkg> && pnpm exec vp test run  # test one package
pnpm exec vp check                          # lint, format & type check
```

Prerequisites: Node.js 18+, pnpm 9+.

## Contribution Workflow

1. **Fork & clone** — fork on GitHub, clone your fork, add `upstream` (`git remote add upstream https://github.com/DemoMacro/docen.git`).
2. **Branch** — branch off `main` (`feat/...`, `fix:...`, `docs/...`, …).
3. **Code** — follow the standards below; match existing style.
4. **Commit** — use [conventional commits](https://www.conventionalcommits.org/): `feat:`, `fix:`, `docs:`, `refactor:`, `perf:`, `test:`, `build:`, `ci:`, `chore:`, `revert:`.
5. **Push & PR** — push to your fork and open a PR against `upstream/main` (checklist below).

## Project Structure

The package map and the file-level tree live in [CLAUDE.md](./CLAUDE.md) → Project / Package Layout — one source, kept current there.

## Coding Standards

### Naming

- **Functions**: camelCase with a semantic prefix — `parse*` / `generate*` (external-format I/O), `resolve*` (DocOpts→JSON), `compile*` (JSON→DocOpts), `project*` (JSON→LayoutDoc), `create*` (factories)
- **Files & directories**: kebab-case
- **Interfaces**: PascalCase, no `I` prefix, `Options` suffix, `readonly` properties
- **Constants**: `as const` objects (not `enum`), SCREAMING_SNAKE_CASE keys, lowercase values

```typescript
export const AlignmentType = {
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
  JUSTIFY: "justify",
} as const;
```

### Loops

| Scenario                        | Use                |
| ------------------------------- | ------------------ |
| Transform into new array        | `.map()`           |
| Filter                          | `.filter()`        |
| Side-effects, async, early exit | `for...of`         |
| Hot paths                       | `for...of` / `for` |

Avoid `.forEach()` — `for...of` is strictly superior.

## Adding DOCX Features

The runtime model is Tiptap JSON; the persistence model is `DocumentOptions` (OOXML). Converters bridge the two; the layout projection turns the same JSON into `LayoutDoc`; the canvas paints it. See CLAUDE.md → Data Model for the projection chain.

### Converter pattern

`resolveDocument` / `compileDocument` (`converters/docx.ts`) walk the tree and assemble `DocumentOptions`. An extension contributes its DOCX expression by scope:

| Scope                      | Extensions                                           | Contribution                                                        |
| -------------------------- | ---------------------------------------------------- | ------------------------------------------------------------------- |
| **Single-node**            | paragraph, heading, image, table, text-style, strike | export `renderDocx(node)` / `parseDocx(opts)` — dispatched per node |
| **Cross-node / container** | blockquote, lists, task-item, mention, details       | export helpers — the converters orchestrate multi-node assembly     |
| **Simple constant**        | page-break, column-break                             | payload inlined in the converter                                    |

### Extension pattern

Custom extensions extend `@tiptap/extension-*` to carry DOCX properties. There is **no DOM rendering path** — `renderHTML` does not exist in this repo:

1. **Attrs** mirror office-open property keys verbatim, with `parseHTML` for clipboard-paste input only
2. **`renderDocx` / `parseDocx`** for DOCX serialization (near-identity passes — attrs are the Options model)
3. Anything visual belongs in the layout projection (`docx/src/layout/project/`) or the painter (`core/src/paint/`), never in the extension

```typescript
export function renderDocx(node: JSONContent): ParagraphOptions {
  /* … */
}
export function parseDocx(opts: ParagraphOptions): Record<string, unknown> {
  /* … */
}

export const Paragraph = BaseParagraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      indent: { default: null, parseHTML: (el) => indentFromElement(el) },
    };
  },
  renderDocx,
  parseDocx,
});
```

### Layout & painting conventions

The rendering pipeline (projection → layout → paint) is described in [CLAUDE.md](./CLAUDE.md) → Architecture: Canvas Rendering. The contribution rules that follow from it:

- **Projection is pure** — reads persisted attrs (+ the style cascade), emits `LayoutDoc` geometry. No Leafer types, no DOM.
- **Painter is dumb** — maps `LayoutDoc` to Leafer elements 1:1. If something renders at the wrong place, fix the projection or the engine, not the painter.
- **Page model** — fixed-height pages and the Word stacking rules are engine semantics (`@docen/layout`) — change them there, once.

## Pull Request Checklist

- [ ] `vp check` passes
- [ ] `pnpm build` + tests succeed for the changed package
- [ ] Naming & patterns follow the standards above
- [ ] Changes are minimal and focused — match existing style
