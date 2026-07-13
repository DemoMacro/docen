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
      "src/presentation.ts",
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
      "@docen/core/image": fileURLToPath(new URL("../core/src/image.ts", import.meta.url)),
      "@docen/core/geometry": fileURLToPath(new URL("../core/src/geometry.ts", import.meta.url)),
      "@docen/core/style": fileURLToPath(new URL("../core/src/style.ts", import.meta.url)),
      "@docen/core/export": fileURLToPath(new URL("../core/src/export.ts", import.meta.url)),
      "@docen/core": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)),
      "@docen/editor": fileURLToPath(new URL("./src/index.ts", import.meta.url)),
    },
  },
  server: {
    open: true,
  },
});
