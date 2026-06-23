import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/ui/**/*",
      "src/document/**/*",
      "src/render/**/*",
      "src/workbook.ts",
      "src/presentation.ts",
    ],
  },
});
