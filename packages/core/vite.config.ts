import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts", "src/image.ts", "src/geometry.ts", "src/style.ts", "src/export.ts"],
  },
});
