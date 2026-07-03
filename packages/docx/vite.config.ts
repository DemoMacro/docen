import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/core.ts",
      "src/editor.ts",
      "src/converters/**/*",
      "src/extensions/**/*",
      // Co-located vitest specs (*.spec.ts) live under src/ for name parity with
      // the module under test, but must not ship in dist (they import vitest).
      "!src/**/*.spec.ts",
      "!src/**/*.test.ts",
    ],
  },
});
