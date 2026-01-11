import { defineBuildConfig } from "unbuild";

export default defineBuildConfig({
  declaration: true,
  entries: ["src/index", "src/types"],
  rollup: {
    emitCJS: true,
    esbuild: {
      minify: true,
    },
    inlineDependencies: true,
  },
});
