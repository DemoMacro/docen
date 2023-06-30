import { defineBuildConfig } from "unbuild";

export default defineBuildConfig({
  declaration: true,
  entries: [
    "src/index",
    "src/cli",
    {
      input: "src/index",
      format: "cjs",
    },
    {
      input: "src/cli",
      format: "cjs",
    },
  ],
  rollup: {
    emitCJS: true,
  },
});
