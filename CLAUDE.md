You are a senior TypeScript developer working on **docen**.

> Coding standards, design patterns, and the contribution workflow live in [CONTRIBUTING.md](./CONTRIBUTING.md). This file is the architectural context an agent must understand before changing code. Read both.

## Project

**docen** is a monorepo for online Office editors.

- **`docen`** — all-in-one aggregate entry: re-exports `@docen/docx` (converters/engine, via `docen/docx`) and `@docen/editor` (`<docen-document>` via `docen/editor`). One dependency covers both headless conversion and the full editor; the root entry stays side-effect-free so converter-only imports remain tree-shakable.
- **`@docen/vue`** — Vue 3 adapter for `@docen/editor`: a typed `<DocenDocument>` component (`v-model` content + `v-slot="{ editor }"` + template-ref expose). `vue` is a peer dependency and `@docen/editor` a regular dependency, so the Vue surface stays isolated from the framework-neutral core.
- **`@docen/editor`** — multi-editor assembly: a Fluent UI host (`<docen-workspace>` + UI surfaces) shared by editor elements `<docen-document>` (today) and `<docen-presentation>`/`<docen-workbook>` (future); all UI surfaces (title-bar/ribbon/status-bar/panes) and engine extensions are contributed by **add-ins** (Office.js-style). Bundles the `@docen/docx` engine; owns the canvas stage, painting, and caret/selection mapping.
- **`@docen/docx`** — the engine: Tiptap DOCX schema + converters + custom extensions + the layout projection (Tiptap JSON → LayoutDoc, incl. WMF/EMF+ metafile replay). No UI.
- **`@docen/layout`** — the layout engine: block/flow/text measurement and pagination producing a paginated `LayoutDoc`. Pure computation, no DOM, no editor types.
- **`@docen/core`** — the scene painter package: LayoutDoc → LeaferJS tree, consumed by the editors' canvas stages. No layout decisions, no editing semantics.
- **`leafer-x-metafile`** — zero-dependency WMF/EMF+ metafile replay into neutral drawing members (no Leafer, no docen types — built to be contributed to the LeaferJS ecosystem as `leafer-x-*`). The docx layout projection consumes it and adapts members into `LayoutDoc`.
- **`@docen/deduplicate`** — document comparison (SimHash + Winnowing fingerprinting, `compareDocuments`/`findDuplicates`) for the editors' future compare feature. Standalone; no editor dependencies.
- **`@office-open/*`** — OOXML parse/generate APIs (external). The canonical document model.

`pptx` and `xlsx` editors are planned — not yet implemented (`packages/editor/src/` has `presentation.ts`/`workbook.ts` stubs). They will reuse the same host + add-in system in `ui/`, swapping only the engine.

## Tech Stack

| Package | Engine                               | Parse/Generate    | Role                                                                                 |
| ------- | ------------------------------------ | ----------------- | ------------------------------------------------------------------------------------ |
| editor  | LeaferJS canvas + Tiptap + Fluent UI | @office-open/docx | Multi-editor host + add-ins + canvas stage/painter/caret-map → `<docen-document>`    |
| docx    | Tiptap (ProseMirror, viewless)       | @office-open/docx | DOCX engine: schema + converters + extensions + layout projection (JSON → LayoutDoc) |
| layout  | —                                    | —                 | Pagination engine: measurement → paginated LayoutDoc                                 |

## Build

- `pnpm install` · `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one) · `vp check` (lint, also via pre-commit hook)
- Build tool: `vp pack` (vite-plus)
- Test: `cd packages/<pkg> && pnpm exec vp test run`

> editor imports `@docen/docx` by package name (→ `dist`), so **docx src changes need `pnpm --filter @docen/docx build`** before they show in the editor demo. editor/src is HMR'd — no build needed.

## Data Model

One document, three projections, each owned by exactly one layer:

| Projection         | Format                                | Owner                       |
| ------------------ | ------------------------------------- | --------------------------- |
| Canonical model    | `DocumentOptions` (@office-open/docx) | file I/O, format conversion |
| Text editing       | Tiptap JSON (DOCX-rich attrs)         | editor transactions         |
| Rendering geometry | `LayoutDoc` (@docen/layout)           | pagination                  |
| Instantiated scene | LeaferJS elements                     | editor canvas painter       |

**Define once, pass through.** office-open's Options types are the single source of truth: Tiptap attrs mirror them verbatim (`renderDocx`/`parseDocx` are near-identity passes), and the layout projection reads the same attrs. No layer re-derives a property another layer already carries; a mapping exists once (stringify side and parse side together).

## API Layering

Standalone functions are core; extension commands are thin wrappers.

```typescript
// Format pipelines — runtime (Tiptap JSON) ↔ external formats
parseDOCX(buffer) → JSONContent                       // DOCX → Tiptap JSON
generateDOCX<T>(json, options?) → Promise<OutputByType[T]>   // prepare + compile + generateDocument
generateDOCXSync<T>(json, packer?) → OutputByType[T]         // sync; no prepare
generateDOCXStream(json, options?) → Promise<ReadableStream>
parseMarkdown / generateMarkdown                       // the second conversion format

// Paste input (input-only; no HTML output exists)
parseHTMLBody(body, schema) → JSONContent             // text/html clipboard → Tiptap JSON

// Model bridge (advanced): resolveDocument (DocOpts→JSON) · compileDocument (JSON→DocOpts)
// · prepareDocument (http img → data URL, in place). Required for http images.
// parseDOCX = parseDocument → resolve → JSON;  generateDOCX = JSON → prepare → compile → generateDocument
```

## Architecture: Canvas Rendering

`<docen-document>` renders through a LeaferJS canvas — **there is no DOM rendering path** (no `renderHTML`, no contenteditable view). The pipeline:

- **Viewless Tiptap editor** (`element: null`): ProseMirror is the editing model only; the EditorView never mounts. Typing/IME goes through a textarea bridge (`document/canvas/edit-bridge.ts`), which also owns clipboard paste (see below).
- **Layout engine** (`@docen/layout`): block/flow/text measurement with Word's stacking rules (docGrid line pitch, snap-to-grid, spacing collapse, table band split with repeated headers and mid-row `cantSplit` handling) produces a paginated `LayoutDoc` of fixed-height pages.
- **Projection** (`docx/src/layout/project.ts`): Tiptap JSON → `LayoutDoc`. This is also where WMF/EMF+ metafiles (`wmf.ts`, `emf-plus.ts`, `wmf-dib.ts`) replay into structured drawing members — vector layers become scene members, not flat bitmaps.
- **Painter** (`editor/document/canvas/scene.ts`): `LayoutDoc` → Leafer elements — the only place the scene is instantiated. `caret-map.ts` maps caret/selection between PM positions and canvas geometry; `stage.ts` owns the Leafer app and zoom.

**Fidelity target:** pixel parity with Word/WPS on real documents, verified page-by-page against PDF exports of the same files. The canvas pipeline (self-drawn layout + paint) is what makes mid-row table splits, vmerge across pages, and docGrid-exact line pitch possible — decoration/contenteditable approaches cannot.

## Architecture: HTML Paste Input

HTML is **input-only**. The extensions' `parseHTML` rules exist for exactly one job: turning pasted styled HTML into document JSON. `parseHTMLBody` (`docx/src/extensions/paste.ts`) is DOM-provider agnostic — the editor passes a native `DOMParser` body, specs pass a linkedom body — and flattens nested `ul`/`ol` before parsing so the ProseMirror parser keeps list nesting levels. There is no HTML generation anywhere.

## Architecture: Add-ins (Office.js-style)

Every editor (`<docen-document>` / `<docen-presentation>` / `<docen-workbook>`) is a **host** (`DocenHost`) whose UI surfaces and engine extensions are contributed by **add-ins** (`DocenAddin`). The default document add-in (`document/addin.ts`) bundles the Office-style ribbon, task panes, commands, and the Tiptap extensions a DOCX editor needs; consumers load extra add-ins to inject ribbon tabs/panes/commands. Implementation in `packages/editor/src/ui/addin/`.

**Naming** aligns to MS Office / Office.js — UI tags use Office terms (`docen-title-bar` / `-ribbon` / `-document-area` / `-status-bar` / `-task-pane` / `-navigation-pane` / `-format-pane`); `RibbonTab` / `Group` / `Control` / `Action` mirror the Office.js manifest. Layer split: `Docx` = file format (`@docen/docx`, `createDocxEditor`); `Document` = editor (`<docen-document>`, `DocumentAddin`). Editor elements self-contain `:host { display:flex; height:100% }` so consumers never add sizing CSS.

## Package Layout

```
packages/docx/src/ — engine + converters + layout projection
  index.ts        Public API
  core.ts         docxExtensions, createDocxEditor
  style-cascade.ts  StylesOptions index/merge (basedOn chains) — shared by resolve/compile/measure
  extensions/     Custom Tiptap extensions (utils.ts, paste.ts, formatting-marks.ts, …)
  converters/     docx.ts (resolveDocument/compileDocument) · styles.ts (quickStyles, effectiveRunProps) · markdown.ts
  layout/         project.ts (JSON → LayoutDoc) · wmf.ts / emf-plus.ts / wmf-dib.ts (metafile replay)

packages/layout/src/ — pagination engine
  block/ flow/ text/   measurement domains
  layout-doc.ts  the LayoutDoc types (rendering projection)
  font.ts        font metrics (incl. CJK)

packages/editor/src/ — multi-editor host + add-ins
  index.ts        Public API (<docen-document> etc.)
  ui/             Shared host + add-in system + Fluent UI surfaces + i18n
    addin/        DocenHost/DocenAddin types · AddinHost base · defineAddin
    components/   ribbon (fast-element) · workspace (title-bar/document-area/status-bar/task-pane/navigation-pane/find-replace/options-dialog/dialog) · context-menu
  document/       <docen-document>
    index.ts      The editor element (open/save/paste, format detection)
    canvas/       stage.ts (Leafer app) · scene.ts (painter) · caret-map.ts (PM pos ↔ canvas geometry) · edit-bridge.ts (textarea + paste)
    addin.ts ribbon.ts commands.ts components/ utils/ extensions/ i18n.ts
  presentation.ts workbook.ts   (future editors — reuse host + add-ins)

packages/core/src/       shared drawing/render helpers (geometry, style, image, export)
packages/deduplicate/src/  document comparison (SimHash + Winnowing)
```

## Performance

- The layout projection re-runs per transaction; metafile replay is fingerprint-cached, media caches key on object identity (not capacity), and repeated lookups are indexed — see the code before adding new per-transaction work.
- Off-screen pages skip painting until visible; only changed subtrees re-project.

## Behavioral Guidelines

- State assumptions explicitly. If uncertain, ask before implementing.
- No features beyond what was asked. No speculative abstractions.
- Touch only what you must. Match existing style.
- Transform tasks into verifiable goals. Loop until verified.
