import swc from "unplugin-swc";
import { defineConfig } from "vite-plus";

export default defineConfig({
  // fast-element's reactive system relies on TS experimental decorators
  // (@customElement / @attr / @observable). Vite's built-in transpiler
  // (esbuild/oxc) does not lower legacy decorators (rolldown#2296), so they
  // reach the browser verbatim → SyntaxError. unplugin-swc takes over .ts
  // transpilation and reads `experimentalDecorators` from tsconfig.json to
  // enable `jsc.transform.legacyDecorator` automatically.
  plugins: [swc.vite()],
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
