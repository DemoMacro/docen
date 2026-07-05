# @docen/editor

![npm version](https://img.shields.io/npm/v/@docen/editor)
![npm downloads](https://img.shields.io/npm/dw/@docen/editor)
![npm license](https://img.shields.io/npm/l/@docen/editor)

> Assembly layer for docen editors — bundles a Fluent UI host with the
> @docen/docx Tiptap engine into turnkey web components like `<docen-document>`,
> and owns Office-style C-route pagination.

## Features

- 🧩 **Turnkey `<docen-document>`** — One custom element bundles the Fluent UI host (title bar, ribbon, document area, status bar, panes, find/replace) with the @docen/docx engine
- 📄 **Office-style pagination** — C-route: fixed-height page boxes with physical overflow reflow, keeping edit == render in a single contenteditable
- 🎨 **Fluent UI surfaces** — Ribbon (buttons, split/toggle buttons, combobox, galleries, color picker), workspace, task/navigation/format panes, context menu
- 🌐 **i18n** — Built-in Chinese (zh-CN) and English (en); add more via `registerTranslation` / `localizationInfo`. Switch live from the status bar (cycles every registered locale) or the Options dialog
- 🌓 **Light/dark theme** — Fluent design tokens drive the chrome; switch via the `theme` attribute
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
filename menu), ribbon, document area,
status bar (page/word count, language indicator, zoom), and panes are all built in.

**Attributes**

Configuration attributes split by reactivity:

- **Reactive** — change at runtime and the component re-renders: `editable`,
  `filename`, `user`, `avatar`, `section-properties`, `styles`, `addins`,
  `theme`.
- **Once** — read only on connect (initial value); runtime control goes through
  methods: `content`, `spellcheck`, `navigation-pane`, `properties-pane`,
  `zoom`, `show-marks`.

The chrome (title bar, ribbon, status bar, panes) is always shown — extend it
via add-ins rather than toggling attributes.

| Attribute            | Default    | Description                                                            |
| -------------------- | ---------- | ---------------------------------------------------------------------- |
| `user`               | —          | Display name shown in the header                                       |
| `avatar`             | —          | Avatar image URL (omitted → initial-letter avatar)                     |
| `filename`           | "Document" | Document name shown in the header and save dialog default              |
| `content`            | —          | Initial document as HTML (once on connect)                             |
| `editable`           | `true`     | `false` makes the surface read-only (reactive)                         |
| `spellcheck`         | `false`    | `true` enables browser spellcheck (perf cost on large docs; once)      |
| `section-properties` | —          | JSON section page setup (size, margins, orientation); reactive         |
| `styles`             | —          | JSON named styles; reactive                                            |
| `addins`             | —          | JSON array of external add-ins (ribbon/task-pane data); see Add-ins    |
| `theme`              | `light`    | `"light" \| "dark"`; drives the Fluent theme                           |
| `navigation-pane`    | —          | `true` opens the navigation (left) pane on connect (once)              |
| `properties-pane`    | —          | `true` opens the properties (right) pane on connect (once)             |
| `zoom`               | `100`      | Initial zoom percent (once); runtime via `setZoom`                     |
| `show-marks`         | `false`    | `true` shows page/section-break markers (once); runtime `setShowMarks` |
| `lang`               | —          | BCP-47 UI locale (`"zh-CN"` / `"en"` / …); per-instance, reactive      |

Unwired ribbon commands (skeleton buttons) render visually but are greyed out
(`disabled`) — the ribbon keeps its full Office shape without dead clicks.

**Methods**

```typescript
class DocenDocument extends HTMLElement {
  // Open — single entry point auto-detects docx/md/html from the extension.
  open(file: File): Promise<void>;
  // Format-specific loaders (use when the format is known up front, e.g. a
  // server-fetched docx buffer with no filename).
  openDOCX(input: File | ArrayBuffer | Uint8Array): Promise<void>;
  openMarkdown(input: File | string): Promise<void>;
  openHTML(input: File | string): Promise<void>;
  saveDOCX(): Promise<Uint8Array>;
  saveMarkdown(): string;
  saveHTML(): string;

  // Runtime model — flat Tiptap JSON (doc > block+). The editor stores pages
  // internally (C-route pagination); getJSON/setJSON unwrap/wrap them so the
  // public model stays page-free (pages must not leak into DOCX export). For
  // Tiptap's own getHTML / getText / setContent / chain, use getEditor().
  getJSON(): JSONContent;
  setJSON(json: JSONContent): void;

  // The underlying @docen/docx Tiptap Editor — the full Tiptap API surface.
  getEditor(): Editor | undefined;
  repaginate(): void;

  // Task-pane visibility (Office.addin.showAsTaskpane / hide equivalent).
  // `id` is "navigation" | "properties"; flips fire docen:taskpane-visibility-change.
  showTaskpane(id: TaskPaneId): void;
  hideTaskpane(id: TaskPaneId): void;
  getTaskpaneState(id: TaskPaneId): boolean;

  // Zoom (Office.Document.zoom.set equivalent; clamped 10–500).
  setZoom(pct: number): void;
  getZoom(): number;

  // Editing/formatting marks (page + section-break markers).
  setShowMarks(on: boolean): void;
  getShowMarks(): boolean;

  // Add-in registry — register/unregister ribbon + command contributions.
  addAddin(addin: DocenAddin): void;
  removeAddin(id: string): void;

  // Office.context.displayLanguage equivalent — read-only current UI locale.
  readonly displayLanguage: string;
}
```

**Events**

All events bubble and compose out of the shadow DOM — listen on the host
element. `docen:save` / `:save-as` / `:open` / `:print` are cancelable: call
`preventDefault()` to take over the action (otherwise the built-in behavior runs).
`docen:save-as` carries `{ format }` (`"docx" | "markdown" | "html"`) — which
Save-As variant the user picked. (`docen:open` is format-agnostic: the host
auto-detects docx/md/html from the chosen file's extension, so it carries no
detail.)

| Event                              | When                                           | Detail                   |
| ---------------------------------- | ---------------------------------------------- | ------------------------ |
| `docen:ready`                      | Editor mounted and ready                       | —                        |
| `docen:change`                     | Document content changed (autosave driver)     | `{ dirty }`              |
| `docen:save`                       | Save button — `preventDefault()` to take over  | —                        |
| `docen:save-as`                    | Save As menu — `preventDefault()` to take over | `{ format }`             |
| `docen:open`                       | Open menu — `preventDefault()` to take over    | —                        |
| `docen:new`                        | New menu — host-only (no built-in action)      | —                        |
| `docen:print`                      | Print menu — `preventDefault()` to take over   | —                        |
| `docen:taskpane-visibility-change` | A task pane opened/closed (method or pane ✕)   | `{ id, visibilityMode }` |
| `docen:zoom-change`                | Zoom changed (button / slider / `setZoom`)     | `{ zoom }`               |
| `docen:marks-change`               | Formatting marks toggled                       | `{ showMarks }`          |
| `docen:lang-change`                | Locale changed (status-bar cycle / Options OK) | `{ lang }`               |

**Slots**

The properties (right) pane body is slot-driven. The default fallback is the
built-in `<docen-format-pane>` (empty state); slot a custom element to take
over the pane — e.g. show image, table, or paragraph properties depending on
the current selection.

| Slot         | Default               | Description                                                 |
| ------------ | --------------------- | ----------------------------------------------------------- |
| `properties` | `<docen-format-pane>` | Right pane body. Slot a component to own the properties UI. |

```html
<docen-document>
  <image-properties slot="properties"></image-properties>
</docen-document>
```

Selection-aware switching is the consumer's responsibility — the Office.js
model leaves task-pane content + navigation to the add-in. Track the selection
(e.g. on `docen:change`) and swap the slotted component (`v-if` / conditional
render). The built-in `<docen-format-pane groups='[{title,fields:[…]}]'>` is a
declarative `radio` / `number` / `color` renderer you can reuse as the slotted
content when declarative fields are enough.

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
  addins='[{"id":"citations","ribbon":[{"tab":"citations","label":"Citations","groups":[{"id":"tools","label":"Tools","controls":[{"type":"button","id":"cite","label":"Cite","event":"bold"}]}]}]}]'
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
      label: "Citations",
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

### Internationalization

The host ships with English (`en`, default) and Chinese (`zh-CN`). Every label
runs through a single `t(key)` lookup, and the locale resolves **per-instance**
from `<docen-document lang>` (forwarded to the internal workspace) — set the
attribute and the ribbon, header, status bar, and Options dialog re-localize
live, with no dependency on `<html lang>`. The status-bar language pill cycles
every registered locale; the Options dialog renders a `<select>` of the same
list.

Add a locale by registering its translation table — the Options dropdown and
the status-bar cycle pick it up with no further wiring:

```typescript
import { registerTranslation } from "@docen/editor";

registerTranslation({
  languageTag: "fr",
  $name: "Français",
  translations: { "ribbon.tab.home": "Accueil" /* … */ },
});
```

Re-registering a tag **merges** (later wins on key conflicts), so an add-in can
extend a built-in locale with its own keys without clobbering the base. The
Office.js manifest shape is supported too — pass `localizationInfo` on an
add-in and the host registers it on `addAddin`:

```typescript
doc.addAddin({
  id: "about",
  localizationInfo: {
    defaultLanguageTag: "en",
    additionalLanguages: [{ languageTag: "zh-CN", translations: { "about.tab": "关于" } }],
  },
  // …ribbon, commands…
});
```

`availableLanguages()` lists every registered tag (for custom pickers); the
read-only `displayLanguage` getter mirrors `Office.context.displayLanguage`.

### UI Bootstrap

```typescript
import { registerComponents, applyTheme } from "@docen/editor";

registerComponents(); // registers <docen-document>
applyTheme("light"); // "light" | "dark"
```

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
