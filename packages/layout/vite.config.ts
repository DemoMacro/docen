import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    // Pretext is patched in-tree (pnpm patch: CJK measure correction, image
    // atoms kept as extra-width items). Bundling it fixes those semantics into
    // dist — a published external import would resolve to the unpatched release.
    deps: { alwaysBundle: ["@chenglou/pretext"] },
  },
});
