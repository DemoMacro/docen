import { defineBasisConfig } from "@funish/basis";

export default defineBasisConfig({
  lint: {
    staged: {
      "*": "pnpm lint",
    },
    project: {
      check:
        "pnpm tsc --noEmit --skipLibCheck && pnpm oxlint --fix --fix-suggestions --type-aware",
      format:
        "pnpm prettier --write --list-different . --plugin=@prettier/plugin-oxc",
    },
    dependencies: {
      allowedLicenses: ["MIT", "ISC", "BSD-2-Clause", "BSD-3-Clause"],
    },
  },
  git: {
    hooks: {
      "pre-commit": "pnpm exec basis lint --staged",
      "commit-msg": "pnpm exec basis git --lint-commit",
    },
    commitMsg: {},
  },
});
