// Nodes
import Blockquote from "@tiptap/extension-blockquote";
import BulletList from "@tiptap/extension-bullet-list";
import CodeBlockLowlight from "@tiptap/extension-code-block-lowlight";
import { Details, DetailsSummary, DetailsContent } from "@tiptap/extension-details";
import Document from "@tiptap/extension-document";
import Emoji from "@tiptap/extension-emoji";
import HardBreak from "@tiptap/extension-hard-break";
import { Heading } from "./extends/heading";
import HorizontalRule from "@tiptap/extension-horizontal-rule";
import { Image } from "./extends/image";
import ListItem from "@tiptap/extension-list-item";
import { Mathematics } from "@tiptap/extension-mathematics";
import OrderedList from "@tiptap/extension-ordered-list";
import { Paragraph } from "./extends/paragraph";
import { Table } from "@tiptap/extension-table";
import TableCell from "@tiptap/extension-table-cell";
import TableHeader from "@tiptap/extension-table-header";
import TableRow from "@tiptap/extension-table-row";
import TaskList from "@tiptap/extension-task-list";
import TaskItem from "@tiptap/extension-task-item";
import Text from "@tiptap/extension-text";

// Marks
import Bold from "@tiptap/extension-bold";
import Code from "@tiptap/extension-code";
import Highlight from "@tiptap/extension-highlight";
import Italic from "@tiptap/extension-italic";
import Link from "@tiptap/extension-link";
import Strike from "@tiptap/extension-strike";
import Subscript from "@tiptap/extension-subscript";
import Superscript from "@tiptap/extension-superscript";
import { TextStyle } from "@tiptap/extension-text-style";
import Underline from "@tiptap/extension-underline";

// Text Style Extensions
import { Color } from "@tiptap/extension-text-style";
import { BackgroundColor } from "@tiptap/extension-text-style";
import { FontFamily } from "@tiptap/extension-text-style";
import { FontSize } from "@tiptap/extension-text-style";
import { LineHeight } from "@tiptap/extension-text-style";

// Nodes
export const tiptapNodeExtensions = [
  Document,
  Paragraph,
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
  Image,
  Mathematics,
  Table,
  TableRow,
  TableCell,
  TableHeader,
  TaskList,
  TaskItem,
  Heading,
];

// Marks
export const tiptapMarkExtensions = [
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
// Nodes
export { Blockquote } from "@tiptap/extension-blockquote";
export { BulletList } from "@tiptap/extension-bullet-list";
export { default as CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";
export { Details, DetailsSummary, DetailsContent } from "@tiptap/extension-details";
export { Document } from "@tiptap/extension-document";
export { Emoji } from "@tiptap/extension-emoji";
export { HardBreak } from "@tiptap/extension-hard-break";
export { Heading } from "./extends/heading";
export { HorizontalRule } from "@tiptap/extension-horizontal-rule";
export { Image } from "./extends/image";
export { ListItem } from "@tiptap/extension-list-item";
export { Mathematics } from "@tiptap/extension-mathematics";
export { OrderedList } from "@tiptap/extension-ordered-list";
export { Paragraph } from "./extends/paragraph";
export { Table } from "@tiptap/extension-table";
export { TableCell } from "@tiptap/extension-table-cell";
export { TableHeader } from "@tiptap/extension-table-header";
export { TableRow } from "@tiptap/extension-table-row";
export { TaskList } from "@tiptap/extension-task-list";
export { TaskItem } from "@tiptap/extension-task-item";
export { Text } from "@tiptap/extension-text";

// Marks
export { Bold } from "@tiptap/extension-bold";
export { Code } from "@tiptap/extension-code";
export { Highlight } from "@tiptap/extension-highlight";
export { Italic } from "@tiptap/extension-italic";
export { Link } from "@tiptap/extension-link";
export { Strike } from "@tiptap/extension-strike";
export { Subscript } from "@tiptap/extension-subscript";
export { Superscript } from "@tiptap/extension-superscript";
export { TextStyle } from "@tiptap/extension-text-style";
export { Underline } from "@tiptap/extension-underline";

// Text Style Extensions
export { Color } from "@tiptap/extension-text-style";
export { BackgroundColor } from "@tiptap/extension-text-style";
export { FontFamily } from "@tiptap/extension-text-style";
export { FontSize } from "@tiptap/extension-text-style";
export { LineHeight } from "@tiptap/extension-text-style";
