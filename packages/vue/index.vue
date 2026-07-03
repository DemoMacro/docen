<script setup lang="ts">
import { ref } from "vue";

// Import from source (not the package name) so vite treats this as workspace
// source, not a pre-bundled dep — mirrors the @docen/editor demo.
import { DocenDocument } from "./src/index.ts";

// v-model: the editor renders this HTML and writes its normalized HTML back
// here on every change — no onReady callback needed.
const content = ref(
  "<h1>Hello from @docen/vue</h1><p>This document is rendered through the Vue adapter.</p>",
);
</script>

<template>
  <DocenDocument v-model="content" filename="Welcome.docx" user="Demo Macro" />
</template>
