import { fileURLToPath, URL } from "node:url";

import vue from "@vitejs/plugin-vue";
import { defineConfig } from "vite-plus";

export default defineConfig({
  // Compiles the index.vue SFC used by the package demo (vp dev).
  plugins: [vue()],
  resolve: {
    alias: {
      // Alias @docen/editor to its workspace source so the demo loads the
      // editor the same way the @docen/editor demo does (source, not dist) —
      // avoids vite pre-bundling the editor dist, which pulls @office-open and
      // its jiti/node:os chain into the browser.
      "@docen/editor": fileURLToPath(new URL("../editor/src/index.ts", import.meta.url)),
    },
  },
  pack: {
    entry: ["src/index.ts"],
  },
});
