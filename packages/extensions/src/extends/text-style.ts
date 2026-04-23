import { TextStyle as BaseTextStyle } from "../tiptap";

export const TextStyle = BaseTextStyle.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      letterSpacing: {
        default: null,
        parseHTML: (element) => element.style.letterSpacing || null,
        renderHTML: (attributes) =>
          attributes.letterSpacing ? { style: `letter-spacing: ${attributes.letterSpacing}` } : {},
      },

      rtl: {
        default: null,
        parseHTML: (element) => (element.dir === "rtl" ? true : null),
        renderHTML: (attributes) => (attributes.rtl ? { dir: "rtl" } : {}),
      },
    };
  },
});
