# @docen/editor

![npm version](https://img.shields.io/npm/v/@docen/editor)
![npm downloads](https://img.shields.io/npm/dw/@docen/editor)
![npm license](https://img.shields.io/npm/l/@docen/editor)

> Assembly layer for docen editors — bundles a Fluent UI host with the
> @docen/docx Tiptap engine into turnkey web components like `<docen-document>`,
> and owns Word-style C-route pagination.

## Features

- 🧩 **Turnkey `<docen-document>`** — One custom element bundles the Fluent UI host (title bar, ribbon, document area, status bar, panes, find/replace) with the @docen/docx engine
- 📄 **Word-style pagination** — C-route: fixed-height page boxes with physical overflow reflow, keeping edit == render in a single contenteditable
- 🎨 **Fluent UI surfaces** — Ribbon (buttons, split/toggle buttons, combobox, galleries, color picker), workspace, task/navigation/format panes, context menu
- 🌐 **i18n** — Built-in Chinese (zh-CN) and English (en); switch live from the user menu
- 🔄 **DOCX round-trip** — Open/save `.docx` through the underlying @docen/docx engine
- 🔌 **Add-ins** — Plug in ribbon tabs, task panes, and commands without touching host internals

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

### Web component: `<docen-document>`

A turnkey WYSIWYG document editor. The title bar (brand, auto-save, save/undo/redo,
Open/Save/Print menu, language menu), ribbon, document area, and panes are all built in.

**Attributes**

Configuration attributes are **reactive** (`observedAttributes`): change them at
runtime and the component re-renders. `content` and `spellcheck` are read once
on connect. The chrome (title bar, ribbon, status bar, panes) is always shown —
extend it via add-ins rather than toggling attributes.

| Attribute            | Default    | Description                                                         |
| -------------------- | ---------- | ------------------------------------------------------------------- |
| `user`               | —          | Display name shown in the header                                    |
| `avatar`             | —          | Avatar image URL (omitted → initial-letter avatar)                  |
| `filename`           | "Document" | Document name shown in the header and save dialog default           |
| `content`            | —          | Initial document as HTML (parsed once on connect)                   |
| `editable`           | `true`     | `false` makes the surface read-only                                 |
| `spellcheck`         | `false`    | `true` enables browser spellcheck (perf cost on large docs)         |
| `section-properties` | —          | JSON section page setup (size, margins, orientation); reactive      |
| `styles`             | —          | JSON named styles; reactive                                         |
| `addins`             | —          | JSON array of external add-ins (ribbon/task-pane data); see Add-ins |

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

| Event           | When                                           | Detail      |
| --------------- | ---------------------------------------------- | ----------- |
| `docen:ready`   | Editor mounted and ready                       | —           |
| `docen:change`  | Document content changed (autosave driver)     | `{ dirty }` |
| `docen:save`    | Save button — `preventDefault()` to take over  | —           |
| `docen:save-as` | Save As menu — `preventDefault()` to take over | —           |
| `docen:open`    | Open menu — `preventDefault()` to take over    | —           |
| `docen:new`     | New menu — host-only (no built-in action)      | —           |
| `docen:print`   | Print menu — `preventDefault()` to take over   | —           |

**Configuration**

The component works out-of-box. Collaborative actions (save, open, print) hand
off to the host via cancelable events; the UI is extended via add-ins.

```html
<!-- Read-only document -->
<docen-document editable="false"></docen-document>
```

```typescript
const doc = document.querySelector<DocenDocument>("#doc")!;

// Take over save (skip the built-in picker → route to your storage)
doc.addEventListener("docen:save", (event) => {
  event.preventDefault();
  saveToStorage(doc.getJSON());
});

// Autosave on change
doc.addEventListener("docen:change", () => scheduleAutosave());
```

**Add-ins**

Plug in ribbon tabs and commands without touching host internals — declaratively
(JSON attribute, data-only — functions can't cross the attribute boundary) or
imperatively (full add-in object via `addAddin`). Task-pane contributions are
reserved for a follow-up (the navigation/format panes are built-in today), so
only ribbon tabs flow through the `addins` attribute right now.

```html
<!-- Declarative: a "Citations" ribbon tab -->
<docen-document
  addins='[{"id":"citations","ribbon":[{"tab":"citations","tabLabel":"Citations","groups":[{"id":"tools","label":"Tools","controls":[{"type":"button","id":"cite","label":"Cite","event":"bold"}]}]}]}]'
></docen-document>
```

```typescript
// Imperative: full add-in (commands / pane-render allowed here)
doc.addAddin({
  id: "citations",
  name: "Citations",
  ribbon: [
    {
      tab: "citations",
      tabLabel: "Citations",
      groups: [
        {
          id: "tools",
          label: "Tools",
          controls: [
            /* … */
          ],
        },
      ],
    },
  ],
});
```

Ribbon control `event` names route to the engine's native Tiptap commands
(`editor.chain().focus().<event>(value).run()`), so a built-in name like `bold`
works directly. Override a command by contributing a Tiptap extension whose
`addCommands` redefines the same name.

### UI Bootstrap

```typescript
import { registerComponents, applyTheme } from "@docen/editor";

registerComponents(); // registers <docen-document>
applyTheme("light"); // "light" | "dark"
```

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
