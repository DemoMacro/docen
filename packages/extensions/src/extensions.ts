// Import all official TipTap extensions from tiptap.ts
import {
  Document,
  Text,
  // Paragraph, // Custom version imported below
  // Heading, // Custom version imported below
  Blockquote,
  HorizontalRule,
  CodeBlockLowlight,
  BulletList,
  OrderedList,
  ListItem,
  TaskList,
  TaskItem,
  Table,
  // TableRow, // Custom version imported below
  TableCell,
  TableHeader,
  // Image, // Custom version imported below
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
  Strike,
  Code,
  Link,
  Highlight,
  Subscript,
  Superscript,
  TextStyle,
  Color,
  BackgroundColor,
  FontFamily,
  FontSize,
  LineHeight,
  TextAlign,
} from "./tiptap";

// Import customized extensions (replace official ones)
import { Heading } from "./extends/heading";
import { Image } from "./extends/image";
import { Paragraph } from "./extends/paragraph";
import { TableRow } from "./extends/table-row";

import { AnyExtension } from "@tiptap/core";

import { all, createLowlight } from "lowlight";

// Nodes
export const tiptapNodeExtensions: AnyExtension[] = [
  Document,
  Paragraph, // Customized version from extends/
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
  }), // Customized version from extends/
  Mathematics,
  Mention,
  Table,
  TableRow, // Customized version from extends/
  TableCell,
  TableHeader,
  TaskList,
  TaskItem,
  Heading, // Customized version from extends/
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

// Export all individual extensions for direct imports
export * from "./tiptap";
export { Heading } from "./extends/heading";
export { Image } from "./extends/image";
export { Paragraph } from "./extends/paragraph";
export { TableRow } from "./extends/table-row";
export { TextAlign } from "@tiptap/extension-text-align";
