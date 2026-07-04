You are a senior TypeScript developer working on **docen**.

> Coding standards, design patterns, and the contribution workflow live in [CONTRIBUTING.md](./CONTRIBUTING.md). This file is the architectural context an agent must understand before changing code. Read both.

## Project

**docen** is a monorepo for online Office editors.

- **`docen`** — all-in-one aggregate entry: re-exports `@docen/docx` (converters/engine, via `docen/docx`) and `@docen/editor` (`<docen-document>` via `docen/editor`). One dependency covers both headless conversion and the full editor; the root entry stays side-effect-free so converter-only imports remain tree-shakable.
- **`@docen/vue`** — Vue 3 adapter for `@docen/editor`: a typed `<DocenDocument>` component (`v-model` content + `v-slot="{ editor }"` + template-ref expose). `vue` is a peer dependency and `@docen/editor` a regular dependency, so the Vue surface stays isolated from the framework-neutral core.
- **`@docen/editor`** — multi-editor assembly: a Fluent UI host (`<docen-workspace>` + UI surfaces) shared by super-components `<docen-document>` (today) and `<docen-presentation>`/`<docen-workbook>` (future); all UI surfaces (title-bar/ribbon/status-bar/panes) and engine extensions are contributed by **add-ins** (Office.js-style). Bundles the `@docen/docx` Tiptap engine for `<docen-document>`; owns C-route pagination.
- **`@docen/docx`** — the engine: Tiptap DOCX editor + converters + custom extensions. No UI.
- **`@office-open/*`** — OOXML parse/generate APIs (external).

`pptx` (LeaferJS) and `xlsx` (RevoGrid) are planned future editors — not yet implemented (`packages/editor/src/` has `presentation.ts`/`workbook.ts` stubs). They will reuse the same host + add-in system in `ui/`, swapping only the engine.

## Tech Stack

| Package | Engine                           | Parse/Generate    | Role                                                                                                 |
| ------- | -------------------------------- | ----------------- | ---------------------------------------------------------------------------------------------------- |
| editor  | Tiptap (ProseMirror) + Fluent UI | @office-open/docx | Multi-editor host + add-ins (Fluent UI surfaces) + docx engine → `<docen-document>`; owns pagination |
| docx    | Tiptap (ProseMirror)             | @office-open/docx | DOCX editor + converters + custom extensions (the engine)                                            |

## Build

- `pnpm install` · `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one) · `vp check` (lint, also via pre-commit hook)
- Build tool: `vp pack` (vite-plus)

> editor imports `@docen/docx` by package name (→ `dist`), so **docx src changes need `pnpm --filter @docen/docx build`** before they show in the editor demo. editor/src is HMR'd — no build needed.

## Data Model

Dual-model — the core mental model:

| Role            | Format                              | Description                                             |
| --------------- | ----------------------------------- | ------------------------------------------------------- |
| **Runtime**     | Tiptap JSON (with DOCX-rich attrs)  | Editor operates on this via transactions                |
| **Persistence** | DocumentOptions (@office-open/docx) | Complete OOXML model for file I/O and format conversion |

Custom Tiptap extensions carry DOCX properties through `parseHTML`/`renderHTML` (node-level), `renderDocx`/`parseDocx`, and `attrs`.

## API Layering

Standalone functions are core; extension commands are thin wrappers.

```typescript
// Format pipelines — runtime (Tiptap JSON) ↔ external formats
parseDOCX(buffer) → JSONContent                       // DOCX → Tiptap JSON
generateDOCX<T>(json, options?) → Promise<OutputByType[T]>   // prepare + compile + generateDocument
generateDOCXSync<T>(json, packer?) → OutputByType[T]         // sync; no prepare
generateDOCXStream(json, options?) → Promise<ReadableStream>
parseHTML / generateHTML / parseMarkdown / generateMarkdown

// Model bridge (advanced): resolveDocument (DocOpts→JSON) · compileDocument (JSON→DocOpts)
// · prepareDocument (http img → data URL, in place). Required for http images.
// parseDOCX = parseDocument → resolve → JSON;  generateDOCX = JSON → prepare → compile → generateDocument
```

## Architecture: Pagination (C-route)

`doc > page+`, each `page` a **fixed-height box**; a ProseMirror `appendTransaction` **physically reflows** overflow to the next page. This is the only route yielding Word-style fixed pages with **edit == render** — a single contenteditable, no separate painter, no selection/cursor coordinate mapping. The page node is editing-time only and never enters DOCX. Implementation rules: CONTRIBUTING.md → Pagination Conventions.

**Why not the alternatives** (researched: Tiptap Pages, docx-editor, ONLYOFFICE, LeaferJS): decoration seams stretch the page on large content; painter dual-rendering splits edit/render and needs coordinate mapping; canvas self-draw means rebuilding a text-layout engine; LeaferJS has no document-level pagination. C-route is the v1 sweet spot.

**Fidelity boundary (vs Word/WPS):** ~90% — fixed pages, overflow reflow, repeated table headers, section geometry, headers/footers, styles, paragraph rules. Not achievable without canvas self-draw: mid-row table split (Word's `cantSplit`; contenteditable can't split a `tr`), vmerge across pages, pixel-exact parity.

## Architecture: Add-ins (Office.js-style)

Every editor (`<docen-document>` / `<docen-presentation>` / `<docen-workbook>`) is a **host** (`DocenHost`) whose UI surfaces and engine extensions are contributed by **add-ins** (`DocenAddin`). The default document add-in (`document/addin.ts`) bundles the Word-style ribbon, task panes, commands, and the Tiptap extensions a DOCX editor needs; consumers load extra add-ins to inject ribbon tabs/panes/commands. Implementation in `packages/editor/src/ui/addin/`.

**Naming** aligns to MS Office / Office.js — UI tags use Office terms (`docen-title-bar` / `-ribbon` / `-document-area` / `-status-bar` / `-task-pane` / `-navigation-pane` / `-format-pane`); `RibbonTab` / `Group` / `Control` / `Action` mirror the Office.js manifest. Layer split: `Docx` = file format (`@docen/docx`, `createDocxEditor`); `Document` = editor (`<docen-document>`, `DocumentAddin`). Super-components self-contain `:host { display:flex; height:100% }` so consumers never add sizing CSS.

## Package Layout

```
packages/docx/src/ — engine + converters
  index.ts        Public API
  core.ts         createDocxEditor(), docxExtensions
  extensions/     Custom Tiptap extensions (utils.ts, formatting-marks.ts, …)
  converters/     docx.ts (DocxManager) · styles.ts (stylesToCss) · html.ts · markdown.ts
  types.ts

packages/editor/src/ — multi-editor host + add-ins
  index.ts        Public API (<docen-document> etc.)
  ui/             Shared host + add-in system + Fluent UI surfaces + i18n
    addin/        DocenHost/DocenAddin types · AddinHost base · defineAddin
    components/   ribbon (fast-element) · workspace (title-bar/document-area/status-bar/task-pane/navigation-pane/format-pane/outline/dialog)
  document/       <docen-document>: index.ts · addin.ts (default document add-in) · ribbon.ts · commands.ts · pagination/ · extensions/
  presentation.ts workbook.ts   (future editors — reuse host + add-ins)
```

## Performance

- Pagination reflow is debounced + cached; only changed pages are re-measured
- `content-visibility: auto` on off-screen pages; disable spellcheck for very large documents

## Behavioral Guidelines

- State assumptions explicitly. If uncertain, ask before implementing.
- No features beyond what was asked. No speculative abstractions.
- Touch only what you must. Match existing style.
- Transform tasks into verifiable goals. Loop until verified.
