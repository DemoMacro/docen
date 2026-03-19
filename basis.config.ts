import { defineBasisConfig } from "@funish/basis/config";

export default defineBasisConfig({
  lint: {
    config: ["--fix", "--fix-suggestions", "--type-aware", "--type-check"],
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
        "*": "pnpm basis lint && pnpm basis fmt",
      },
    },
  },
});
