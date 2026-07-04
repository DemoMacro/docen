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
});
