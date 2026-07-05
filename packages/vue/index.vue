<script setup lang="ts">
import type { JSONContent } from "@docen/docx";
import { ref } from "vue";

// Import from source (not the package name) so vite treats this as workspace
// source, not a pre-bundled dep — mirrors the @docen/editor demo.
import { DocenDocument } from "./src/index.ts";

// v-model: Tiptap JSON (page nodes unwrapped). The editor injects this via
// host.setJSON on ready and writes host.getJSON() back here (debounced) on
// change — preserves DOCX-rich attrs that an HTML round-trip would drop.
const content = ref<JSONContent>({
  type: "doc",
  content: [
    {
      type: "heading",
      attrs: { level: 1 },
      content: [{ type: "text", text: "Hello from @docen/vue" }],
    },
    {
      type: "paragraph",
      content: [{ type: "text", text: "This document is rendered through the Vue adapter." }],
    },
  ],
});
</script>

<template>
  <DocenDocument v-model="content" filename="Welcome.docx" user="Demo Macro" />
</template>
