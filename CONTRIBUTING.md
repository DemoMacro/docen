# Contributing to docen

Thanks for contributing! This guide covers the **workflow** for contributing and the **coding standards** that keep docen consistent. For architectural context (data models, API layering, design decisions), see [CLAUDE.md](./CLAUDE.md).

## Development Setup

```bash
pnpm install                          # install dependencies
pnpm build                            # build all packages
cd packages/<pkg> && pnpm build       # build one package
vp check                              # lint & format
```

Prerequisites: Node.js 18+, pnpm 9+.

## Contribution Workflow

1. **Fork & clone** — fork on GitHub, clone your fork, add `upstream` (`git remote add upstream https://github.com/DemoMacro/docen.git`).
2. **Branch** — branch off `main` (`feat/...`, `fix:...`, `docs/...`, …).
3. **Code** — follow the standards below; match existing style.
4. **Verify** — `vp check` passes; `pnpm build` succeeds for the changed package.
5. **Commit** — use [conventional commits](https://www.conventionalcommits.org/): `feat:`, `fix:`, `docs:`, `refactor:`, `perf:`, `test:`, `build:`, `ci:`, `chore:`, `revert:`.
6. **Push & PR** — push to your fork and open a PR against `upstream/main`.

## Project Structure

```
packages/
  editor/   @docen/editor  (assembly: Fluent UI shell + @docen/docx → <docen-document>; owns pagination)
  docx/     @docen/docx    (Tiptap DOCX editor + converters + custom extensions)
```

- **@docen/editor** — assembly layer: Fluent UI shell (under `src/ui/`) + docx engine.
- **@docen/docx** — engine + converters, no UI.

See CLAUDE.md → Package Layout for the file-level tree.

## Coding Standards

### Naming

- **Functions**: camelCase with a semantic prefix — `parse*` / `generate*` (external-format I/O), `resolve*` (DocOpts→JSON), `compile*` (JSON→DocOpts), `create*` (factories)
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

The runtime model is Tiptap JSON; the persistence model is `DocumentOptions` (OOXML). Converters bridge the two. See CLAUDE.md → Data Model & API Layering for the data flow.

### Converter pattern

`DocxManager` (`converters/docx.ts`) walks the tree and assembles `DocumentOptions`. An extension contributes its DOCX expression by scope:

| Scope                      | Extensions                                           | Contribution                                                        |
| -------------------------- | ---------------------------------------------------- | ------------------------------------------------------------------- |
| **Single-node**            | paragraph, heading, image, table, text-style, strike | export `renderDocx(node)` / `parseDocx(opts)` — dispatched per node |
| **Cross-node / container** | blockquote, lists, task-item, mention, details       | export helpers — `DocxManager` orchestrates multi-node assembly     |
| **Simple constant**        | page-break, column-break                             | payload inlined in `DocxManager`                                    |

### Extension pattern

Custom extensions extend `@tiptap/extension-*` to carry DOCX properties:

1. **Attrs** with `parseHTML` only (no attribute-level renderHTML for nodes)
2. **Node-level `renderHTML`** computes all CSS at once (avoids style-merge conflicts)
3. **`renderDocx` / `parseDocx`** for DOCX serialization (single-node only)

Mark extensions (text-style, strike) keep attribute-level `renderHTML`.

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
      indent: { default: null, parseHTML: (el) => el.style.marginLeft || null },
    };
  },
  renderHTML({ node, HTMLAttributes }) {
    const styles = renderParagraphStyles(node.attrs);
    const attrs = styles.length ? { ...HTMLAttributes, style: styles.join(";") } : HTMLAttributes;
    return ["p", attrs, 0] as const;
  },
  renderDocx,
  parseDocx,
});
```

### Pagination conventions (C-route)

`doc > page+`, fixed-height page boxes, physical reflow. See CLAUDE.md → Pagination for the architecture.

- **Page node is round-trip transparent** — never enters DOCX. `DocxManager` operates on flat `doc > block+`; the page node exists only at the editor layer. Do NOT add page-node handling to `DocxManager`. (`pageBreak`/`sectionBreak` ARE semantic nodes that round-trip.)
- **Fixed page box** — `.docen-page { height: <content area>; overflow: hidden }`. Use `height`, not `min-height` (min-height lets content stretch the page).
- **Reflow** — break at block boundaries first (whole paragraph), then whole table rows; never mid-glyph. Binary-search the break. Debounce + cache measurements (DOM `offsetHeight` is ground truth).
- **Paragraph rules** (Word defaults) — widow/orphan control, keepNext (heading + next block), keepLines.
- **Table across pages** — whole-row move; clone `tableHeader` on continuation pages; clip + warn for over-tall rows (no infinite loop). Mid-row split is out of scope (see CLAUDE.md → Fidelity boundary).

## Pull Request Checklist

- [ ] `vp check` passes
- [ ] `pnpm build` succeeds for the changed package
- [ ] Naming & patterns follow the standards above
- [ ] Changes are minimal and focused — match existing style
