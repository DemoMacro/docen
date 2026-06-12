import { defineBasisConfig } from "@funish/basis/config";

export default defineBasisConfig({
  release: {
    npm: {
      additionalTag: "edge",
    },
  },
});
