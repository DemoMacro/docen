import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    // Two public entries mirroring upstream's "." and "./rich-inline" exports;
    // the shared modules (analysis/measurement/line-break/bidi) emit as chunks.
    entry: ["src/layout.ts", "src/rich-inline.ts"],
  },
});
