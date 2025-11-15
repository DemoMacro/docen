// Nodes
import Blockquote from "@tiptap/extension-blockquote";
import BulletList from "@tiptap/extension-bullet-list";
import CodeBlockLowlight from "@tiptap/extension-code-block-lowlight";
import {
  Details,
  DetailsSummary,
  DetailsContent,
} from "@tiptap/extension-details";
import Document from "@tiptap/extension-document";
import Emoji from "@tiptap/extension-emoji";
import HardBreak from "@tiptap/extension-hard-break";
import Heading from "@tiptap/extension-heading";
import HorizontalRule from "@tiptap/extension-horizontal-rule";
import Image from "@tiptap/extension-image";
import ListItem from "@tiptap/extension-list-item";
import { Mathematics } from "@tiptap/extension-mathematics";
import OrderedList from "@tiptap/extension-ordered-list";
import Paragraph from "@tiptap/extension-paragraph";
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

// Tiptap extensions used for schema generation

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
];

// Complete extension set
export const tiptapExtensions = [
  ...tiptapNodeExtensions,
  ...tiptapMarkExtensions,
];
