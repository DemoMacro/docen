import { Table as BaseTable } from "../tiptap";

export const Table = BaseTable.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      layout: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },

      alignment: {
        default: null,
        parseHTML: (element) => {
          const ml = element.style.marginLeft;
          const mr = element.style.marginRight;
          if (ml === "auto" && mr === "auto") return "center";
          if (ml === "auto") return "right";
          if (mr === "auto") return "left";
          return null;
        },
        renderHTML: (attributes) => {
          if (!attributes.alignment) return {};
          switch (attributes.alignment) {
            case "center":
              return { style: "margin-left: auto; margin-right: auto" };
            case "right":
              return { style: "margin-left: auto; margin-right: 0" };
            default:
              return { style: "margin-left: 0; margin-right: auto" };
          }
        },
      },

      cellSpacing: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
    };
  },
});
