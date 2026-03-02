import { defineBasisConfig } from "@funish/basis/config";

export default defineBasisConfig({
  lint: {
    config: ["--fix", "--fix-suggestions", "--type-aware"],
  },
  fmt: {
    config: ["--write", "."],
  },
  publish: {
    npm: {
      additionalTag: "edge",
    },
  },
  git: {
    hooks: {
      "pre-commit": "pnpm basis git staged",
      "commit-msg": "pnpm basis git lint-commit",
    },
    staged: {
      rules: {
        "*.{ts,tsx,js,jsx}": "basis lint --fix",
        "*.{json,md,yml,yaml}": "basis fmt --write",
      },
    },
  },
});
