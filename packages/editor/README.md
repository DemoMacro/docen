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

await doc.openDOCX(file); // open from a File
doc.openDOCXFromBuffer(arrayBuffer); // ...or from an ArrayBuffer/Uint8Array
const output = await doc.saveDOCX(); // → Uint8Array
```

## API

### Super-component: `<docen-document>`

A turnkey WYSIWYG document editor. The header (brand, auto-save, save/undo/redo,
Open/Save/Print menu, language menu), ribbon, canvas, and panes are all built in.

**Attributes**

| Attribute  | Description                                               |
| ---------- | --------------------------------------------------------- |
| `user`     | Display name shown in the header                          |
| `filename` | Document name shown in the header and save dialog default |

**Methods**

```typescript
class DocenDocument extends HTMLElement {
  // File I/O
  openDOCX(file: File): Promise<void>;
  openDOCXFromBuffer(buffer: ArrayBuffer | Uint8Array): void;
  saveDOCX(): Promise<Uint8Array>;

  // Tiptap JSON (the runtime model)
  getJSON(): JSONContent;
  setJSON(json: JSONContent): void;

  // Underlying engine + pagination
  getEditor(): Editor | undefined; // the @docen/docx Tiptap Editor
  repaginate(): void;
}
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
