import { TableHeader as BaseTableHeader } from "../tiptap";

export const TableHeader = BaseTableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      noWrap: {
        default: null,
        parseHTML: (element) => (element.style.whiteSpace === "nowrap" ? true : null),
        renderHTML: (attributes) => (attributes.noWrap ? { style: "white-space: nowrap" } : {}),
      },

      textDirection: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
    };
  },
});
