# @docen/vue

![npm version](https://img.shields.io/npm/v/@docen/vue)
![npm downloads](https://img.shields.io/npm/dw/@docen/vue)
![npm license](https://img.shields.io/npm/l/@docen/vue)

> Vue 3 adapter for [`@docen/editor`](../editor) — a typed `<DocenDocument>` component: `v-model` for content, a `v-slot="{ editor }"` scope, and a template-ref expose.

The editor UI is a framework-neutral web component (`<docen-document>`); this package wraps it for idiomatic Vue — typed props/emits, two-way content binding, and direct access to the underlying Tiptap editor.

## Install

```bash
pnpm add @docen/vue
```

`@docen/editor` comes along as a dependency; `vue` is a peer dependency (install it in your app).

## Usage

```vue
<script setup lang="ts">
import { ref } from "vue";
import type { JSONContent } from "@docen/docx";
import { DocenDocument } from "@docen/vue";
import { parseDOCX } from "@docen/docx";

// v-model carries Tiptap JSON (page nodes unwrapped); the template ref
// exposes the Tiptap editor plus a getJSON/setJSON pair.
const content = ref<JSONContent>({ type: "doc", content: [{ type: "paragraph" }] });
const editorRef = ref();

async function open(file: File) {
  const json = await parseDOCX(await file.arrayBuffer());
  // setJSON routes through the host loader so doc.attrs.styles survive —
  // editor.commands.setContent would drop them.
  editorRef.value?.setJSON(json);
}
</script>

<template>
  <DocenDocument ref="editorRef" v-model="content" filename="Welcome.docx" editable />
</template>
```

Reaching the editor without a ref — via the default slot scope:

```vue
<DocenDocument v-model="content" v-slot="{ editor }">
  <button :disabled="!editor" @click="editor?.chain().focus().toggleBold().run()">B</button>
</DocenDocument>
```

## v-model

`v-model` binds the document content as Tiptap JSON (page nodes unwrapped — the public model is a flat `doc > block+`; the editor repackages it into C-route pages internally):

- Setting `modelValue` calls `host.setJSON(json)`, which routes through the host's loader (a fresh `EditorState`) so document-level attrs (`styles`, `core`, `sectionProperties`) survive — `editor.commands.setContent` would drop them. The set is skipped when `modelValue` is the same reference the adapter just emitted (round-trip echo break).
- Every editor change emits `update:modelValue` with `host.getJSON()`, debounced by the `debounce` prop (default 300 ms; 0 = synchronous). A DOCX import triggers many change events as pagination reflows — one `getJSON` per quiet window instead of one per transaction.

The initial `modelValue` seeds the editor on `docen:ready` via `host.setJSON` (not the `content` attribute, which would re-serialize a large string and lose doc-level attrs).

## Props

Mirror the `<docen-document>` attributes. Pass `undefined` to leave an attribute unset (the web component's default applies).

| Prop                | Type    | Attribute            | Notes                                       |
| ------------------- | ------- | -------------------- | ------------------------------------------- |
| `modelValue`        | object  | — (setJSON)          | Tiptap JSON (page nodes unwrapped); two-way |
| `debounce`          | number  | —                    | Emit debounce ms (default 300; 0 = sync)    |
| `filename`          | string  | `filename`           |                                             |
| `editable`          | boolean | `editable`           |                                             |
| `spellcheck`        | boolean | `spellcheck`         |                                             |
| `user` / `avatar`   | string  | `user` / `avatar`    | Identity in the header                      |
| `sectionProperties` | object  | `section-properties` | JSON page setup (size/margin/orientation)   |
| `styles`            | object  | `styles`             | JSON named-styles model                     |
| `addins`            | array   | `addins`             | JSON external add-ins (ribbon/task-pane)    |
| `theme`             | string  | `theme`              | Fluent built-in key; reactive               |
| `navigationPane`    | boolean | `navigation-pane`    | Initial nav-pane visibility (once)          |
| `propertiesPane`    | boolean | `properties-pane`    | Initial properties-pane visibility (once)   |
| `zoom`              | number  | `zoom`               | Initial zoom percent (once)                 |
| `showMarks`         | boolean | `show-marks`         | Initial marks visibility (once)             |
| `lang`              | string  | `lang`               | BCP-47 UI locale; per-instance, reactive    |

## Events

Re-emitted from the web component's `docen:*` events. Host events forward the
**native `CustomEvent`** — read the payload from `e.detail`:

- `update:modelValue` — editor content changed (drives v-model; Tiptap JSON)
- `@change` — `e.detail: { dirty: true }`; fires on every content-changing transaction (selection-only tx skipped). Lightweight notification only — for the actual content use `v-model` (`update:modelValue` is debounced); don't call `getJSON()` in this handler on large docs (it's O(n)).
- `@zoom-change` (`{ zoom }`), `@taskpane-visibility-change` (`{ id, visibilityMode }`), `@marks-change` (`{ showMarks }`)
- `@lang-change` (`{ lang }`), `@theme-change` (`{ theme }`)

### Taking over save / open / print

`@save`, `@save-as`, `@open`, `@new`, `@print` are **cancelable**: call
`e.preventDefault()` to suppress the host's built-in action (save → native DOCX
save-as dialog, open → file picker, print → browser print) and take over:

```vue
<script setup lang="ts">
import type { DocenSaveEvent } from "@docen/vue";

async function onSave(e: DocenSaveEvent) {
  e.preventDefault(); // stop the host's built-in save-as dialog
  // ...persist to your backend instead
}
</script>

<template>
  <DocenDocument @save="onSave" />
</template>
```

## Template ref

The ref exposes `{ editor, getElement(), getDisplayLanguage(), getJSON(), setJSON(json), addAddin(addin) }`, where `editor` is the Tiptap `Editor` (undefined until the editor is live). `getJSON()/setJSON()` mirror the host's page-unwrapping loaders; `getDisplayLanguage()` returns the current UI locale (`Office.context.displayLanguage` equivalent).

`addAddin(addin)` registers an Office.js-style add-in at runtime (commands / task-pane / ribbon / mini-toolbar). Use it for add-ins carrying functions — the `addins` prop only accepts JSON-serializable data (attribute values are strings), so command handlers and pane render functions must register through the ref. Wait for the host to be live (the exposed `editor` is non-undefined) before calling:

```vue
<script setup lang="ts">
import { ref, watch } from "vue";
import { DocenDocument, type DocenAddin } from "@docen/vue";

const docRef = ref();
const myAddin: DocenAddin = {
  id: "my-addin",
  commands: { "my-action": () => alert("run") },
};
// addAddin needs the host live — `editor` is undefined until docen:ready.
watch(
  () => docRef.value?.editor,
  (ed) => ed && docRef.value?.addAddin(myAddin),
);
</script>

<template>
  <DocenDocument ref="docRef" />
</template>
```

## Internationalization

The adapter re-exports the i18n API from `@docen/editor`, so a Vue app registers
locales from the same entry. The host ships with English (`en`, default) and
Chinese (`zh-CN`); add more by registering a translation table:

```typescript
import { registerTranslation } from "@docen/vue";

registerTranslation({
  languageTag: "fr",
  $name: "Français",
  translations: { "ribbon.tab.home": "Accueil" /* … */ },
});
```

Bind the locale with `:lang` + `@lang-change` — the host forwards the attribute
to its workspace and every label re-resolves live:

```vue
<script setup lang="ts">
import { ref } from "vue";
import { DocenDocument } from "@docen/vue";

const lang = ref("en");
</script>

<template>
  <DocenDocument :lang="lang" @lang-change="lang = $event.lang" />
</template>
```

Re-registering a tag merges, so an add-in can extend a built-in locale with its
own keys. `availableLanguages()` lists every registered tag.

## Theming

`<docen-document>` ships 8 Fluent built-in themes. Bind with `:theme` +
`@theme-change` (mirrors `:lang` / `@lang-change`):

```vue
<script setup lang="ts">
import { ref } from "vue";
import { DocenDocument } from "@docen/vue";

const theme = ref("light");
</script>

<template>
  <DocenDocument :theme="theme" @theme-change="theme = $event.theme" />
</template>
```

Keys: `light`, `dark`, `high-contrast`, `teams-light`, `teams-dark`,
`teams-high-contrast`, `teams-light-v21`, `teams-dark-v21`. In dark / high
contrast the page paper and body ink both follow the theme (Word Dark Mode
behavior); documents still print with their original colors.

Register a custom brand theme with `registerTheme(id, createLightTheme(brand))`
from `@docen/editor` — `brand` is a full 16-step `BrandVariants` ramp (10 darkest
→ 160 lightest; missing shades make `createLightTheme` emit `undefined` tokens
that crash `setGlobalTheme`) — then set `:theme="id"`. The built-in set is
iterated via `builtinThemes.keys()`.

## Why a separate package?

The adapter imports `@docen/editor`, which registers the `<docen-document>` custom element as a side effect. Isolating that in `@docen/vue` (peer-depending on `vue`) keeps non-Vue consumers of `@docen/editor` unaffected, and keeps Vue out of the framework-neutral core.

### SSR (Nuxt)

That side effect runs `customElements.define(...)` at module load. Node has no `customElements` global, so a server bundle that evaluates `import "@docen/vue"` throws a `ReferenceError`. `<ClientOnly>` only guards **rendering**, not module evaluation — use one of:

- render the page client-only: `definePageMeta({ ssr: false })`, or name the component `*.client.vue`;
- or load the adapter dynamically inside the client tree:

```ts
import { defineAsyncComponent } from "vue";
const DocenDocument = defineAsyncComponent(() => import("@docen/vue").then((m) => m.DocenDocument));
```

## License

MIT © [Demo Macro](https://www.demomacro.com/)
