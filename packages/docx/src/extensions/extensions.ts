import { all, createLowlight } from "lowlight";

import { Extension, type AnyExtension } from "../core";
import { Blockquote } from "./blockquote";
import { CodeBlock } from "./code-block";
import { ColumnBreak } from "./column-break";
import { Details, DetailsSummary, DetailsContent } from "./details";
import { Document } from "./document";
import { Heading } from "./heading";
import { Image } from "./image";
import { Mention } from "./mention";
import { OrderedList } from "./ordered-list";
import { PageBreak } from "./page-break";
import { Paragraph } from "./paragraph";
import { Passthrough } from "./passthrough";
import { Strike } from "./strike";
import { Table } from "./table";
import { TableCell } from "./table-cell";
import { TableHeader } from "./table-header";
import { TableRow } from "./table-row";
import { TaskItem } from "./task-item";
import { TextStyle } from "./text-style";
import {
  Text,
  HorizontalRule,
  CodeBlockLowlight,
  BulletList,
  ListItem,
  ListKeymap,
  TaskList,
  HardBreak,
  Emoji,
  Mathematics,
  Bold,
  Italic,
  Underline,
  Code,
  Link,
  Highlight,
  Subscript,
  Superscript,
  TextAlign,
  Dropcursor,
  Gapcursor,
  TrailingNode,
  UndoRedo,
} from "./tiptap";

// Nodes
export const tiptapNodeExtensions: AnyExtension[] = [
  Document,
  Paragraph,
  Text,
  HardBreak,
  PageBreak,
  ColumnBreak,
  Passthrough,
  Blockquote,
  OrderedList,
  BulletList,
  ListItem,
  CodeBlock.configure({
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
];

// Complete extension set
export const docxExtensions: AnyExtension[] = [...tiptapNodeExtensions, ...tiptapMarkExtensions];

// DocxKit options type
export interface DocxKitOptions {
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

export const DocxKit = Extension.create<DocxKitOptions>({
  name: "docxKit",

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
export { Document } from "./document";
export { Heading } from "./heading";
export { Image } from "./image";
export { Passthrough } from "./passthrough";
export { Paragraph } from "./paragraph";
export { TableRow } from "./table-row";
export { Table } from "./table";
export { TableCell } from "./table-cell";
export { TableHeader } from "./table-header";
export { Strike } from "./strike";
export { TextStyle } from "./text-style";
export { TextAlign } from "@tiptap/extension-text-align";
