import { defineBuildConfig } from "@funish/basis/config";

export default defineBuildConfig({
  entries: [
    {
      entry: ["src/index", "src/docx"],
      inlineOnly: false,
    },
  ],
});
