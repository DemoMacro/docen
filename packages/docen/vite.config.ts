import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts", "src/docx.ts", "src/editor.ts"],
  },
});
