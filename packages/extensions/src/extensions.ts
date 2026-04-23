import {
  Document,
  Text,
  Blockquote,
  HorizontalRule,
  CodeBlockLowlight,
  BulletList,
  OrderedList,
  ListItem,
  ListKeymap,
  TaskList,
  TaskItem,
  HardBreak,
  Details,
  DetailsSummary,
  DetailsContent,
  Emoji,
  Mention,
  Mathematics,
  Bold,
  Italic,
  Underline,
  Code,
  Link,
  Highlight,
  Subscript,
  Superscript,
  Color,
  BackgroundColor,
  FontFamily,
  FontSize,
  LineHeight,
  TextAlign,
  Dropcursor,
  Gapcursor,
  TrailingNode,
  UndoRedo,
} from "./tiptap";

import { Heading } from "./extends/heading";
import { Image } from "./extends/image";
import { Paragraph } from "./extends/paragraph";
import { TableRow } from "./extends/table-row";
import { Table } from "./extends/table";
import { TableCell } from "./extends/table-cell";
import { TableHeader } from "./extends/table-header";
import { Strike } from "./extends/strike";
import { TextStyle } from "./extends/text-style";

import { Extension, type AnyExtension } from "@tiptap/core";

import { all, createLowlight } from "lowlight";

// Nodes
export const tiptapNodeExtensions: AnyExtension[] = [
  Document,
  Paragraph,
  Text,
  HardBreak,
  Blockquote,
  OrderedList,
  BulletList,
  ListItem,
  CodeBlockLowlight.configure({
    lowlight: createLowlight(all),
  }),
  Details,
  DetailsSummary,
  DetailsContent,
  Emoji,
  HorizontalRule,
  Image.configure({
    inline: true,
  }),
  Mathematics,
  Mention,
  Table,
  TableRow,
  TableCell,
  TableHeader,
  TaskList,
  TaskItem,
  Heading,
  TextAlign.configure({
    types: ["heading", "paragraph"],
  }),
];

// Marks
export const tiptapMarkExtensions: AnyExtension[] = [
  Bold,
  Code,
  Highlight,
  Italic,
  Link,
  Strike,
  Subscript,
  Superscript,
  TextStyle,
  Underline,
  Color,
  BackgroundColor,
  FontFamily,
  FontSize,
  LineHeight,
];

// Complete extension set
export const tiptapExtensions: AnyExtension[] = [...tiptapNodeExtensions, ...tiptapMarkExtensions];

// StarterKit options type
export interface StarterKitOptions {
  bold?: Record<string, any> | false;
  blockquote?: Record<string, any> | false;
  bulletList?: Record<string, any> | false;
  code?: Record<string, any> | false;
  codeBlock?: Record<string, any> | false;
  document?: false;
  dropcursor?: Record<string, any> | false;
  gapcursor?: false;
  hardBreak?: Record<string, any> | false;
  heading?: Record<string, any> | false;
  undoRedo?: Record<string, any> | false;
  horizontalRule?: Record<string, any> | false;
  italic?: Record<string, any> | false;
  listItem?: Record<string, any> | false;
  listKeymap?: Record<string, any> | false;
  link?: Record<string, any> | false;
  orderedList?: Record<string, any> | false;
  paragraph?: Record<string, any> | false;
  strike?: Record<string, any> | false;
  text?: false;
  underline?: Record<string, any> | false;
  trailingNode?: Record<string, any> | false;
}

export const StarterKit = Extension.create<StarterKitOptions>({
  name: "docenKit",

  addExtensions() {
    const extensions: AnyExtension[] = [];

    if (this.options.bold !== false) {
      extensions.push(Bold.configure(this.options.bold));
    }
    if (this.options.blockquote !== false) {
      extensions.push(Blockquote.configure(this.options.blockquote));
    }
    if (this.options.bulletList !== false) {
      extensions.push(BulletList.configure(this.options.bulletList));
    }
    if (this.options.code !== false) {
      extensions.push(Code.configure(this.options.code));
    }
    if (this.options.codeBlock !== false) {
      extensions.push(
        CodeBlockLowlight.configure({
          lowlight: createLowlight(all),
          ...this.options.codeBlock,
        }),
      );
    }
    if (this.options.document !== false) {
      extensions.push(Document);
    }
    if (this.options.dropcursor !== false) {
      extensions.push(Dropcursor.configure(this.options.dropcursor));
    }
    if (this.options.gapcursor !== false) {
      extensions.push(Gapcursor);
    }
    if (this.options.hardBreak !== false) {
      extensions.push(HardBreak.configure(this.options.hardBreak));
    }
    if (this.options.heading !== false) {
      extensions.push(Heading.configure(this.options.heading));
    }
    if (this.options.undoRedo !== false) {
      extensions.push(UndoRedo.configure(this.options.undoRedo));
    }
    if (this.options.horizontalRule !== false) {
      extensions.push(HorizontalRule.configure(this.options.horizontalRule));
    }
    if (this.options.italic !== false) {
      extensions.push(Italic.configure(this.options.italic));
    }
    if (this.options.listItem !== false) {
      extensions.push(ListItem.configure(this.options.listItem));
    }
    if (this.options.listKeymap !== false) {
      extensions.push(ListKeymap.configure(this.options.listKeymap));
    }
    if (this.options.link !== false) {
      extensions.push(Link.configure(this.options.link));
    }
    if (this.options.orderedList !== false) {
      extensions.push(OrderedList.configure(this.options.orderedList));
    }
    if (this.options.paragraph !== false) {
      extensions.push(Paragraph.configure(this.options.paragraph));
    }
    if (this.options.strike !== false) {
      extensions.push(Strike.configure(this.options.strike));
    }
    if (this.options.text !== false) {
      extensions.push(Text);
    }
    if (this.options.underline !== false) {
      extensions.push(Underline.configure(this.options.underline));
    }
    if (this.options.trailingNode !== false) {
      extensions.push(TrailingNode.configure(this.options.trailingNode));
    }

    return extensions;
  },
});

// Export all individual extensions for direct imports
export * from "./tiptap";
export { Heading } from "./extends/heading";
export { Image } from "./extends/image";
export { Paragraph } from "./extends/paragraph";
export { TableRow } from "./extends/table-row";
export { Table } from "./extends/table";
export { TableCell } from "./extends/table-cell";
export { TableHeader } from "./extends/table-header";
export { Strike } from "./extends/strike";
export { TextStyle } from "./extends/text-style";
export { TextAlign } from "@tiptap/extension-text-align";
