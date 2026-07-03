import { defineConfig } from "vite-plus";

export default defineConfig({
  test: {
    // Scope tests to the workspace packages so temp checkouts under .temp/
    // (other repos staged for reference/e2e) aren't picked up by the default
    // glob and reported as failures.
    include: ["packages/**/*.{spec,test}.ts"],
    benchmark: {
      reporters: ["default"],
    },
    sequence: {
      concurrent: true,
    },
  },
  fmt: {
    sortImports: {
      type: "natural",
    },
    sortPackageJson: true,
    sortTailwindcss: {},
  },
  lint: {
    options: {
      typeAware: true,
      typeCheck: true,
    },
  },
  staged: {
    "*": "vp check --fix",
  },
});
