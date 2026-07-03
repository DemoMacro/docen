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
import { DocenDocument } from "@docen/vue";
import { parseDOCX } from "@docen/docx";

// v-model keeps content in sync; the template ref exposes the Tiptap editor.
const content = ref("<p>Hello</p>");
const editorRef = ref();

async function open(file: File) {
  const json = await parseDOCX(await file.arrayBuffer());
  editorRef.value?.editor?.commands.setContent(json);
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

`v-model` binds the document content as HTML:

- Setting `modelValue` calls `editor.commands.setContent(html)` (skipped when the editor already holds that HTML).
- Every editor change emits `update:modelValue` with `editor.getHTML()`.

The initial value also seeds the editor on connect (via the `content` attribute).

## Props

Mirror the `<docen-document>` attributes. Pass `undefined` to leave an attribute unset (the web component's default applies).

| Prop                | Type    | Attribute            | Notes                                     |
| ------------------- | ------- | -------------------- | ----------------------------------------- |
| `modelValue`        | string  | `content` (v-model)  | Document HTML; two-way                    |
| `filename`          | string  | `filename`           |                                           |
| `editable`          | boolean | `editable`           |                                           |
| `spellcheck`        | boolean | `spellcheck`         |                                           |
| `toolbar`           | boolean | `toolbar`            | Show/hide the ribbon                      |
| `header`            | boolean | `header`             | Show/hide the app header                  |
| `statusBar`         | boolean | `status-bar`         | Show/hide the status bar                  |
| `navigationPane`    | string  | `navigation-pane`    | `open` / `closed` / `hidden`              |
| `propertiesPane`    | string  | `properties-pane`    | `open` / `closed` / `hidden`              |
| `tabs`              | string  | `tabs`               | Comma list, e.g. `"home,review"`          |
| `closable`          | boolean | `closable`           | Render the close (×) button               |
| `user` / `avatar`   | string  | `user` / `avatar`    | Identity in the header                    |
| `sectionProperties` | object  | `section-properties` | JSON page setup (size/margin/orientation) |
| `styles`            | object  | `styles`             | JSON named-styles model                   |

## Events

Re-emitted from the web component's `docen:*` events:

- `update:modelValue` — editor content changed (drives v-model)
- `@change`, `@save`, `@save-as`, `@open`, `@new`, `@print`, `@request-close`

## Template ref

The ref exposes `{ editor, getElement() }`, where `editor` is the Tiptap `Editor` (undefined until the editor is live).

## Why a separate package?

The adapter imports `@docen/editor`, which registers the `<docen-document>` custom element as a side effect. Isolating that in `@docen/vue` (peer-depending on `vue`) keeps non-Vue consumers of `@docen/editor` unaffected, and keeps Vue out of the framework-neutral core.

## License

MIT © [Demo Macro](https://www.demomacro.com/)
