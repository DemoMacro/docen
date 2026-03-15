import { defineBuildConfig } from "@funish/basis/config";

export default defineBuildConfig({
  entries: [
    {
      entry: ["src/index"],
      deps: {
        onlyBundle: false,
      },
    },
  ],
});
