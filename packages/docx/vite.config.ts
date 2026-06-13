import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts", "src/core.ts", "src/editor.ts", "src/converters/*", "src/extensions/*"],
  },
});
