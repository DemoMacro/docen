import { Node as TiptapNode } from "@tiptap/core";
import { CodeBlockLowlight } from "@tiptap/extension-code-block-lowlight";
import { Emoji } from "@tiptap/extension-emoji";
import { HorizontalRule } from "@tiptap/extension-horizontal-rule";
import { Mathematics } from "@tiptap/extension-mathematics";
import { all, createLowlight } from "lowlight";

import { Extension, type AnyExtension } from "../core";
import { Blockquote } from "./blockquote";
import { CodeBlock } from "./code-block";
import { ColumnBreak } from "./column-break";
import { Details, DetailsSummary, DetailsContent } from "./details";
import { Document } from "./document";
import { FormattingMarks } from "./formatting-marks";
import { Image } from "./image";
import { Link } from "./link";
import { Bold, Code, Highlight, Italic, Strike, Subscript, Superscript, Underline } from "./marks";
import { Mention } from "./mention";
import { PageBreak } from "./page-break";
import { Paragraph } from "./paragraph";
import { Passthrough, InlinePassthrough } from "./passthrough";
import { SectionBreak } from "./section-break";
import { Tab } from "./tab";
import { Table } from "./table";
import { TableCell } from "./table-cell";
import { TableRow } from "./table-row";
import { TextStyle } from "./text-style";
import { TocField } from "./toc-field";
import { Insertion, Deletion } from "./track-change";
import { WpgGroup } from "./wpg-group";
import { WpsShape } from "./wps-shape";

// Nodes
/** The inline text atom — plain `Node.create` (the upstream extension added
 *  nothing we use). */
const Text = TiptapNode.create({ name: "text", group: "inline" });

/** Hard line break (<w:br/>) — `Node.create` inline atom. The upstream
 *  extension's keymap/input rules never fire in the viewless canvas route
 *  (typing goes through the textarea bridge). */
const HardBreak = TiptapNode.create({
  name: "hardBreak",
  inline: true,
  group: "inline",
  selectable: false,
  parseHTML() {
    return [{ tag: "br" }];
  },
  renderHTML() {
    return ["br"] as const;
  },
});

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
];

// DocxKit options type
export interface DocxKitOptions {
  bold?: Record<string, any> | false;
  blockquote?: Record<string, any> | false;
  code?: Record<string, any> | false;
  codeBlock?: Record<string, any> | false;
  document?: false;
  hardBreak?: Record<string, any> | false;
  horizontalRule?: Record<string, any> | false;
  italic?: Record<string, any> | false;
  link?: Record<string, any> | false;
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
      extensions.push(Bold);
    }
    if (this.options.blockquote !== false) {
      extensions.push(Blockquote.configure(this.options.blockquote));
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
      extensions.push(HardBreak);
    }
    if (this.options.horizontalRule !== false) {
      extensions.push(HorizontalRule.configure(this.options.horizontalRule));
    }
    if (this.options.italic !== false) {
      extensions.push(Italic);
    }
    if (this.options.link !== false) {
      extensions.push(Link.configure(this.options.link));
    }
    if (this.options.paragraph !== false) {
      extensions.push(Paragraph.configure(this.options.paragraph));
    }
    if (this.options.strike !== false) {
      extensions.push(Strike);
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
export { HorizontalRule } from "@tiptap/extension-horizontal-rule";
export { Mathematics } from "@tiptap/extension-mathematics";
export { Bold, Code, Highlight, Italic, Strike, Subscript, Superscript, Underline } from "./marks";
export { Document, createDocument } from "./document";
export { Paragraph } from "./paragraph";
export { detectHeadingLevel } from "./paragraph";
// Flat list model: generated numbering references + level builders shared by
// compile (definition registration) and the editor list commands.
export {
  BULLET_GLYPHS,
  BULLET_REFERENCE,
  ORDERED_FORMATS,
  ORDERED_REFERENCE_PREFIX,
  nextOrderedReference,
} from "./list-numbering";
export { Blockquote } from "./blockquote";
export { CodeBlock } from "./code-block";
export { ColumnBreak } from "./column-break";
export { SectionBreak } from "./section-break";
export { Details, DetailsSummary, DetailsContent } from "./details";
export { Mention } from "./mention";
export { Table } from "./table";
export { TableRow } from "./table-row";
export { TableCell } from "./table-cell";
export { Image } from "./image";
export { Link } from "./link";
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
