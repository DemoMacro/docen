import { Paragraph as BaseParagraph } from "../tiptap";
import { renderBorderCSS } from "../utils";

export const Paragraph = BaseParagraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      indentLeft: {
        default: null,
        parseHTML: (element) => element.style.marginLeft || null,
        renderHTML: (attributes) =>
          attributes.indentLeft ? { style: `margin-left: ${attributes.indentLeft}` } : {},
      },

      indentRight: {
        default: null,
        parseHTML: (element) => element.style.marginRight || null,
        renderHTML: (attributes) =>
          attributes.indentRight ? { style: `margin-right: ${attributes.indentRight}` } : {},
      },

      indentFirstLine: {
        default: null,
        parseHTML: (element) => element.style.textIndent || null,
        renderHTML: (attributes) =>
          attributes.indentFirstLine ? { style: `text-indent: ${attributes.indentFirstLine}` } : {},
      },

      indentFirstLineChars: {
        default: null,
        parseHTML: () => null,
        renderHTML: (attributes) => {
          if (attributes.indentFirstLineChars == null) return {};
          const em = attributes.indentFirstLineChars / 100;
          return { style: `text-indent: ${em}em` };
        },
      },

      spacingBefore: {
        default: null,
        parseHTML: (element) => element.style.marginTop || null,
        renderHTML: (attributes) =>
          attributes.spacingBefore ? { style: `margin-top: ${attributes.spacingBefore}` } : {},
      },

      spacingAfter: {
        default: null,
        parseHTML: (element) => element.style.marginBottom || null,
        renderHTML: (attributes) =>
          attributes.spacingAfter ? { style: `margin-bottom: ${attributes.spacingAfter}` } : {},
      },

      shading: {
        default: null,
        parseHTML: (element) => {
          const bg = element.style.backgroundColor;
          if (!bg) return null;
          const fill = bg.startsWith("#") ? bg : `#${bg}`;
          return { fill, type: "clear" };
        },
        renderHTML: (attributes) => {
          if (!attributes.shading?.fill) return {};
          return { style: `background-color: ${attributes.shading.fill}` };
        },
      },

      borderTop: {
        default: null,
        parseHTML: () => null,
        renderHTML: (attributes) => {
          const css = renderBorderCSS(attributes.borderTop);
          return css ? { style: `border-top: ${css}` } : {};
        },
      },
      borderBottom: {
        default: null,
        parseHTML: () => null,
        renderHTML: (attributes) => {
          const css = renderBorderCSS(attributes.borderBottom);
          return css ? { style: `border-bottom: ${css}` } : {};
        },
      },
      borderLeft: {
        default: null,
        parseHTML: () => null,
        renderHTML: (attributes) => {
          const css = renderBorderCSS(attributes.borderLeft);
          return css ? { style: `border-left: ${css}` } : {};
        },
      },
      borderRight: {
        default: null,
        parseHTML: () => null,
        renderHTML: (attributes) => {
          const css = renderBorderCSS(attributes.borderRight);
          return css ? { style: `border-right: ${css}` } : {};
        },
      },
    };
  },
});
