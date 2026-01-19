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
} from "./tiptap";

// Import customized extensions (replace official ones)
import { Heading } from "./extends/heading";
import { Image } from "./extends/image";
import { Paragraph } from "./extends/paragraph";
import { TableRow } from "./extends/table-row";

import { Extensions } from "@tiptap/core";

// Nodes
export const tiptapNodeExtensions: Extensions = [
  Document,
  Paragraph, // Customized version from extends/
  Text,
  HardBreak,
  Blockquote,
  OrderedList,
  BulletList,
  ListItem,
  CodeBlockLowlight,
  Details,
  DetailsSummary,
  DetailsContent,
  Emoji,
  HorizontalRule,
  Image, // Customized version from extends/
  Mathematics,
  Mention,
  Table,
  TableRow, // Customized version from extends/
  TableCell,
  TableHeader,
  TaskList,
  TaskItem,
  Heading, // Customized version from extends/
];

// Marks
export const tiptapMarkExtensions: Extensions = [
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
export const tiptapExtensions = [...tiptapNodeExtensions, ...tiptapMarkExtensions];

// Export all individual extensions for direct imports
export * from "./tiptap";
export { Heading } from "./extends/heading";
export { Image } from "./extends/image";
export { Paragraph } from "./extends/paragraph";
export { TableRow } from "./extends/table-row";
