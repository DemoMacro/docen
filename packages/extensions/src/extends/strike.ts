import { Strike as BaseStrike } from "../tiptap";

export const Strike = BaseStrike.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      doubleStrike: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
    };
  },
});
