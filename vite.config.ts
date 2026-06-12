import { defineConfig } from "vite-plus";

export default defineConfig({
  test: {
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
