# @docen/editor

![npm version](https://img.shields.io/npm/v/@docen/editor)
![npm license](https://img.shields.io/npm/l/@docen/editor)

> Assembly layer for docen editors — bundles a Fluent UI shell with the
> @docen/docx Tiptap engine into turnkey web-component super-components like
> `<docen-document>`, and owns Word-style C-route pagination.

## Features

- 🧩 **Turnkey `<docen-document>`** — One custom element bundles the Fluent UI shell (ribbon, canvas, panes, find/replace) with the @docen/docx engine
- 📄 **Word-style pagination** — C-route: fixed-height page boxes with physical overflow reflow, keeping edit == render in a single contenteditable
- 🎨 **Fluent UI shell** — Ribbon (buttons, split/toggle buttons, combobox, galleries, color picker), workspace, task/properties/navigation panes, context menu
- 🌐 **i18n** — Built-in Chinese (zh-CN) and English (en); switch live from the user menu
- 🔄 **DOCX round-trip** — Open/save `.docx` through the underlying @docen/docx engine
- 🔌 **Stub super-components** — `<docen-presentation>` and `<docen-workbook>` reserved for future LeaferJS/RevoGrid editors

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/editor

# Install with npm
$ npm install @docen/editor
```

## Quick Start

`<docen-document>` is a self-contained custom element — register the components,
apply a theme, and drop it in.

```html
<docen-document id="doc" user="Demo Macro" filename="Welcome.docx"></docen-document>

<script type="module">
  import { registerComponents, applyTheme } from "@docen/editor";

  registerComponents(); // register all custom elements
  applyTheme("light"); // "light" | "dark"
</script>
```

Open and save DOCX imperatively:

```typescript
const doc = document.querySelector<DocenDocument>("#doc")!;

await doc.openDOCX(file); // File | ArrayBuffer | Uint8Array
const output = await doc.saveDOCX(); // → Uint8Array
```

## API

### Super-component: `<docen-document>`

A turnkey WYSIWYG document editor. The header (brand, auto-save, save/undo/redo,
Open/Save/Print menu, language menu), ribbon, canvas, and panes are all built in.

**Attributes**

Chrome-visibility and configuration attributes are **reactive** (`observedAttributes`):
change them at runtime and the component re-renders. `content` and `spellcheck`
are read once on connect.

| Attribute         | Default    | Description                                                          |
| ----------------- | ---------- | -------------------------------------------------------------------- |
| `user`            | —          | Display name shown in the header                                     |
| `avatar`          | —          | Avatar image URL (omitted → initial-letter avatar)                   |
| `filename`        | "Document" | Document name shown in the header and save dialog default            |
| `content`         | —          | Initial document as HTML (parsed once on connect)                    |
| `editable`        | `true`     | `false` makes the surface read-only                                  |
| `spellcheck`      | `false`    | `true` enables browser spellcheck (perf cost on large docs)          |
| `toolbar`         | `true`     | `false` hides the ribbon                                             |
| `tabs`            | all        | Comma-separated ribbon tab whitelist, e.g. `tabs="home,review,view"` |
| `header`          | `true`     | `false` hides the app header                                         |
| `navigation-pane` | `open`     | `open` \| `closed` \| `hidden` — left navigation pane state          |
| `properties-pane` | `open`     | `open` \| `closed` \| `hidden` — right properties pane state         |
| `status-bar`      | `true`     | `false` hides the footer status bar                                  |
| `closable`        | —          | Renders a close (×) button that emits `docen:request-close`          |

Unwired ribbon commands (skeleton buttons) render visually but are greyed out
(`disabled`) — the ribbon keeps its full Office shape without dead clicks.

**Methods**

```typescript
class DocenDocument extends HTMLElement {
  // DOCX I/O
  openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void>;
  saveDOCX(): Promise<Uint8Array>;

  // Runtime model — flat Tiptap JSON (doc > block+). The editor stores pages
  // internally (C-route pagination); getJSON/setJSON unwrap/wrap them so the
  // public model stays page-free (pages must not leak into DOCX export). For
  // Tiptap's own getHTML / getText / setContent / chain, use getEditor().
  getJSON(): JSONContent;
  setJSON(json: JSONContent): void;

  // The underlying @docen/docx Tiptap Editor — the full Tiptap API surface.
  getEditor(): Editor | undefined;
  repaginate(): void;
}
```

**Events**

All events bubble and compose out of the shadow DOM — listen on the host
element. `docen:save` / `:save-as` / `:open` / `:print` are cancelable: call
`preventDefault()` to take over the action (otherwise the built-in behavior runs).

| Event                 | When                                           | Detail      |
| --------------------- | ---------------------------------------------- | ----------- |
| `docen:ready`         | Editor mounted and ready                       | —           |
| `docen:change`        | Document content changed (autosave driver)     | `{ dirty }` |
| `docen:request-close` | User clicked × (only rendered with `closable`) | —           |
| `docen:save`          | Save button — `preventDefault()` to take over  | —           |
| `docen:save-as`       | Save As menu — `preventDefault()` to take over | —           |
| `docen:open`          | Open menu — `preventDefault()` to take over    | —           |
| `docen:new`           | New menu — host-only (no built-in action)      | —           |
| `docen:print`         | Print menu — `preventDefault()` to take over   | —           |

**Configuration**

The component works out-of-box, but every part of the chrome is toggleable and
collaborative actions hand off to the host.

```html
<!-- Read-only, minimal chrome, two tabs -->
<docen-document
  editable="false"
  header="false"
  status-bar="false"
  tabs="home,review"
></docen-document>
```

```typescript
const doc = document.querySelector<DocenDocument>("#doc")!;

// Take over save (skip the built-in picker → route to your storage)
doc.addEventListener("docen:save", (event) => {
  event.preventDefault();
  saveToStorage(doc.getJSON());
});

// The × button only appears with `closable`; the component never unmounts
// itself — the host decides what "close" means.
doc.setAttribute("closable", "");
doc.addEventListener("docen:request-close", () => doc.remove());

// Autosave on change
doc.addEventListener("docen:change", () => scheduleAutosave());

// Reactive: change tabs at runtime (observedAttributes)
doc.setAttribute("tabs", "home,insert,view");
```

### UI Bootstrap

```typescript
import { registerComponents, applyTheme } from "@docen/editor";

registerComponents(); // registers <docen-document> (+ the stubs)
applyTheme("light"); // "light" | "dark"
```

### Stub Super-components

`<docen-presentation>` (LeaferJS) and `<docen-workbook>` (RevoGrid) are reserved
for future editors — not yet implemented.

## Architecture

```
@docen/editor (assembly)
  <docen-document> super-component
    ├── Fluent UI shell (ribbon, workspace, panes, find/replace)
    ├── @docen/docx engine (Tiptap + DOCX/HTML/Markdown converters)
    └── C-route pagination (fixed-height pages, appendTransaction reflow)
          page-node · page-plugin · measure · paragraph-split · table-split
```

- **Assembly layer** — bundles the shell + engine into one custom element and owns pagination (which lives at the editor layer and never enters DOCX)
- **C-route pagination** — `doc > page+`, fixed-height page boxes, ProseMirror `appendTransaction` physically reflows overflow, so edit == render
- **Engine reuse** — all DOCX parse/generate flows through `@docen/docx` standalone functions

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
