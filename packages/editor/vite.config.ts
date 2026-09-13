import { fileURLToPath, URL } from "node:url";

import { defineConfig } from "vite-plus";

export default defineConfig({
  // fast-element's reactive system relies on TS experimental decorators
  // (@customElement / @attr / @observable). Vite+'s oxc transform now lowers
  // legacy decorators automatically (oxc-project/oxc#4047; the earlier gap
  // rolldown/rolldown#2296 has been resolved), reading `experimentalDecorators`
  // from tsconfig.json. tsconfig has no `emitDecoratorMetadata`, so oxc's
  // partial-metadata caveat does not apply.
  pack: {
    entry: [
      "src/index.ts",
      "src/ui/**/*",
      "src/document/**/*",
      "src/workbook.ts",
      "src/presentation/**/*",
      // Co-located vitest specs live under src/ but must not ship in dist
      // (they import vitest and would resurface as stale suites in test runs).
      "!src/**/*.spec.ts",
      "!src/**/*.test.ts",
    ],
  },
  resolve: {
    alias: {
      // `pnpm demo` serves the demos in /demo from this package's source so edits
      // HMR instantly — the demos import `@docen/editor` / `@docen/core` by
      // package name, and these aliases point them at workspace source instead
      // of pre-bundling dist (which would pull @office-open + jiti/node:os).
      // Each subpath export maps to its own source entry. @docen/docx is NOT
      // aliased: editor's source imports it by package name → dist, so docx src
      // changes still need `pnpm --filter @docen/docx build`.
      // The geometry subpath must alias BEFORE the bare package: alias keys
      // prefix-match, and the bare entry would otherwise rewrite
      // "@docen/core/geometry" into "<entry>/geometry", a path that resolves
      // to nothing.
      "@docen/core/geometry": fileURLToPath(
        new URL("../core/src/geometry/index.ts", import.meta.url),
      ),
      "@docen/core": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)),
      // The layout engine serves from source too: the canvas demo imports it
      // directly, and @docen/docx's dist (not aliased) bare-imports it from
      // its layout/ subpath — this alias resolves both without pre-bundling.
      "@docen/layout": fileURLToPath(new URL("../layout/src/index.ts", import.meta.url)),
      "@docen/pptx": fileURLToPath(new URL("../pptx/src/index.ts", import.meta.url)),
      "@docen/editor": fileURLToPath(new URL("./src/index.ts", import.meta.url)),
    },
  },
  server: {
    open: true,
  },
});
