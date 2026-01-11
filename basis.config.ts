import { defineBasisConfig } from "@funish/basis";

export default defineBasisConfig({
  lint: {
    staged: {
      "*": "bun lint",
    },
    project: {
      check: "oxlint --fix --fix-suggestions --type-aware",
      format: "oxfmt --write . --ignore-path .gitignore",
    },
  },
  git: {
    hooks: {
      "pre-commit": "basis lint --staged",
      "commit-msg": "basis git --lint-commit",
    },
  },
});
