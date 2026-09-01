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
      // Co-located vitest specs live under src/ but must not ship in dist
      // (they import vitest and would resurface as stale suites in test runs).
      "!src/**/*.spec.ts",
      "!src/**/*.test.ts",
    ],
    // Pretext is patched in-tree (pnpm patch: CJK measure correction, image
    // atoms kept as extra-width items). Bundling it fixes those semantics into
    // dist — a published external import would resolve to the unpatched release.
    deps: { alwaysBundle: ["@chenglou/pretext"] },
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
      "@docen/core": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)),
      // The layout engine serves from source too: the canvas demo imports it
      // directly, and @docen/docx's dist (not aliased) bare-imports it from
      // its layout/ subpath — this alias resolves both without pre-bundling.
      "@docen/layout": fileURLToPath(new URL("../layout/src/index.ts", import.meta.url)),
      "@docen/editor": fileURLToPath(new URL("./src/index.ts", import.meta.url)),
    },
  },
  server: {
    open: true,
  },
});
