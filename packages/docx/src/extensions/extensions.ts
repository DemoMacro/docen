import { CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";
import { Emoji } from "@tiptap/extension-emoji";
import { HardBreak } from "@tiptap/extension-hard-break";
import { HorizontalRule } from "@tiptap/extension-horizontal-rule";
import { ListItem } from "@tiptap/extension-list-item";
import { Mathematics } from "@tiptap/extension-mathematics";
import { TaskList } from "@tiptap/extension-task-list";
import { Text } from "@tiptap/extension-text";
import { TextAlign } from "@tiptap/extension-text-align";
import { all, createLowlight } from "lowlight";

import { Extension, type AnyExtension } from "../core";
import { Blockquote } from "./blockquote";
import { BulletList } from "./bullet-list";
import { CodeBlock } from "./code-block";
import { ColumnBreak } from "./column-break";
import { Details, DetailsSummary, DetailsContent } from "./details";
import { Document } from "./document";
import { FormattingMarks } from "./formatting-marks";
import { Heading } from "./heading";
import { Image } from "./image";
import { Link } from "./link";
import { ListAggregator } from "./list-aggregator";
import { Bold, Code, Highlight, Italic, Subscript, Superscript, Underline } from "./marks";
import { Mention } from "./mention";
import { OrderedList } from "./ordered-list";
import { PageBreak } from "./page-break";
import { Paragraph } from "./paragraph";
import { Passthrough, InlinePassthrough } from "./passthrough";
import { SectionBreak } from "./section-break";
import { Strike } from "./strike";
import { Tab } from "./tab";
import { Table } from "./table";
import { TableCell } from "./table-cell";
import { TableHeader } from "./table-header";
import { TableRow } from "./table-row";
import { TaskItem } from "./task-item";
import { TextStyle } from "./text-style";
import { TocField } from "./toc-field";
import { Insertion, Deletion } from "./track-change";
import { WpgGroup } from "./wpg-group";
import { WpsShape } from "./wps-shape";

// Nodes
export const tiptapNodeExtensions: AnyExtension[] = [
  Document,
  Paragraph,
  Text,
  HardBreak,
  PageBreak,
  ColumnBreak,
  Tab,
  SectionBreak,
  Passthrough,
  InlinePassthrough,
  TocField,
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
  WpgGroup,
  WpsShape,
  // NOTE: Mathematics (blockMath/inlineMath) renders via KaTeX in the editor but
  // has no DOCX conversion yet — DOCX compile drops math content. latex↔OMML
  // conversion is separate work (office-open has OMML parse/stringify via its
  // MathInput type, but no latex bridge). Kept registered so the editor works.
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
  Deletion,
  Highlight,
  Insertion,
  Italic,
  Link,
  Strike,
  Subscript,
  Superscript,
  TextStyle,
  Underline,
];

// DOCX schema + DOCX-specific extensions. Editing-behavior extensions
// (UndoRedo/Dropcursor/Gapcursor/TrailingNode/ListKeymap/CharacterCount/Focus)
// live in @docen/editor — the engine stays free of editing-UX concerns.
// Converters (html/markdown) use this array as schema; those extensions add no
// schema, so omitting them does not affect conversion.
export const docxExtensions: AnyExtension[] = [
  ...tiptapNodeExtensions,
  ...tiptapMarkExtensions,
  FormattingMarks,
  // Plain Extensions adding no schema — ListAggregator owns the DOCX → Tiptap
  // list-tree rebuild (declares parseDocxAggregator). Registered so DocxManager
  // collects its rule; it never reaches the ProseMirror schema.
  ListAggregator,
];

// DocxKit options type
export interface DocxKitOptions {
  bold?: Record<string, any> | false;
  blockquote?: Record<string, any> | false;
  bulletList?: Record<string, any> | false;
  code?: Record<string, any> | false;
  codeBlock?: Record<string, any> | false;
  document?: false;
  hardBreak?: Record<string, any> | false;
  heading?: Record<string, any> | false;
  horizontalRule?: Record<string, any> | false;
  italic?: Record<string, any> | false;
  listItem?: Record<string, any> | false;
  link?: Record<string, any> | false;
  orderedList?: Record<string, any> | false;
  paragraph?: Record<string, any> | false;
  strike?: Record<string, any> | false;
  text?: false;
  underline?: Record<string, any> | false;
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
    if (this.options.hardBreak !== false) {
      extensions.push(HardBreak.configure(this.options.hardBreak));
    }
    if (this.options.heading !== false) {
      extensions.push(Heading.configure(this.options.heading));
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

    return extensions;
  },
});

// Export all individual extensions for direct imports from @docen/docx.
// Re-export explicitly (no `export *`) so the public surface is visible.
// Customized extensions export their local version; upstream-only ones re-export
// from @tiptap/* directly, base marks (with DOCX hooks) from ./marks.
export { CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";
export { Emoji } from "@tiptap/extension-emoji";
export { HardBreak } from "@tiptap/extension-hard-break";
export { HorizontalRule } from "@tiptap/extension-horizontal-rule";
export { ListItem } from "@tiptap/extension-list-item";
export { Mathematics } from "@tiptap/extension-mathematics";
export { TaskList } from "@tiptap/extension-task-list";
export { Text } from "@tiptap/extension-text";
export { TextAlign } from "@tiptap/extension-text-align";
export { Bold, Code, Highlight, Italic, Subscript, Superscript, Underline } from "./marks";
export { Document, createDocument } from "./document";
export { Paragraph } from "./paragraph";
export { Heading } from "./heading";
export { Blockquote } from "./blockquote";
export { BulletList } from "./bullet-list";
export { OrderedList } from "./ordered-list";
export { CodeBlock } from "./code-block";
export { ColumnBreak } from "./column-break";
export { SectionBreak } from "./section-break";
export { Details, DetailsSummary, DetailsContent } from "./details";
export { TaskItem } from "./task-item";
export { Mention } from "./mention";
export { Table } from "./table";
export { TableRow } from "./table-row";
export { TableCell } from "./table-cell";
export { TableHeader } from "./table-header";
export { Image } from "./image";
export { Link } from "./link";
export { Strike } from "./strike";
export { TextStyle } from "./text-style";
export { Insertion, Deletion } from "./track-change";
export { FormattingMarks } from "./formatting-marks";
export { PageBreak } from "./page-break";
export {
  WpgGroup,
  wpsShapeStyles,
  type WpsShapeStyles,
  type WpsShapeStandalone,
} from "./wpg-group";
export { WpsShape } from "./wps-shape";
export { Passthrough, InlinePassthrough } from "./passthrough";
export { TocField } from "./toc-field";
export { Tab } from "./tab";
export { scrollCaretToTop, scrollContainerOf } from "./scroll";
