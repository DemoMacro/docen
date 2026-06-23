/**
 * @docen/docx — Type definitions.
 *
 * Design: Tiptap attr interfaces mirror @office-open/docx Option interfaces
 * via the AttrNullable mapped type. This maximizes reuse and keeps attr
 * structure identical to the persistence model, so parseDocx/renderDocx
 * become near-identity mappings.
 *
 * Attr design principles:
 * - Mirror @office-open/docx Option interfaces, INCLUDING nesting
 *   (indent/spacing/border/run/frame are nested objects)
 * - Store office-open native values (twips, points, AlignmentType strings)
 *   so DOCX round-trip is lossless by construction
 *   NOTE: `size` is in POINTS (new office-open convention), not half-points
 * - CSS conversion happens only in renderHTML via utils mappers
 * - Keep Tiptap structural names where base extensions require them
 *   (colspan, rowspan, colwidth, src, alt, level)
 *
 * @module
 */

import type {
  ParagraphPropertiesOptionsBase,
  RunStylePropertiesOptions,
  TableOptions,
  TableRowPropertiesOptionsBase,
  TableCellOptions,
  Floating,
  ImageOptions,
} from "@office-open/docx";
import type { JSONContent as TiptapJSONContent } from "@tiptap/core";

// ============================================================
// Layer 1: Re-export @office-open/docx types (persistence model)
// ============================================================

// The five office-open types used internally for attr derivation (below) are
// imported once and re-exported by reference — instead of a second
// `export ... from "@office-open/docx"` — so each type has a single
// dependency declaration in this file.
export type { TiptapJSONContent as JSONContent };
export type {
  ParagraphPropertiesOptionsBase,
  RunStylePropertiesOptions,
  TableOptions,
  TableRowPropertiesOptionsBase,
  TableCellOptions,
};

export type {
  // Document structure
  DocumentOptions,
  SectionOptions,
  SectionChild,
  SectionPropertiesOptions,
  // Paragraph
  ParagraphOptions,
  ParagraphChild,
  ParagraphPropertiesOptions,
  ParagraphStylePropertiesOptions,
  LevelParagraphStylePropertiesOptions,
  // Run
  RunOptions,
  RunPropertiesOptions,
  ParagraphRunPropertiesOptions,
  // Image
  ImageOptions,
  ImageChild,
  MediaTransformation,
  // Table
  TableRowOptions,
  // Indent, spacing, borders, shading
  IndentAttributesProperties,
  SpacingProperties,
  BordersOptions,
  BorderOptions,
  ShadingAttributesProperties,
  // Table structural types (reused in attr interfaces)
  TableBordersOptions,
  TableCellBordersOptions,
  TableFloatOptions,
  TableLookOptions,
  TableWidthProperties,
  TableLayoutType,
  TableVerticalAlign,
  HeightRule,
  WidthType,
  Margins,
  // Run structural types (reused in attr interfaces)
  FontAttributesProperties,
  EmphasisMarkType,
  UnderlineType,
  // Alignment, heading, tab stops, line rule
  AlignmentType,
  HeadingLevel,
  TabStopDefinition,
  TabStopType,
  TabStopPosition,
  LeaderType,
  LineRuleType,
  TextAlignmentType,
  // Floating, hyperlinks, math
  Floating,
  ExternalHyperlinkOptions,
  InternalHyperlinkOptions,
  MathChild,
  MathInput,
  // Underline, highlight
  HighlightColor,
  // Frame (paragraph text frames)
  FrameOptions,
  // Bookmark, ruby
  BookmarkOptions,
  RubyOptions,
} from "@office-open/docx";

// ============================================================
// Layer 2: Tiptap attr interfaces (runtime model)
//
// Derived from @office-open/docx Option interfaces via AttrNullable.
// Every property becomes `T | null` (required, explicit null) to match
// ProseMirror's attr storage model (every declared attr is stored,
// even when null). Nesting reduces the attr count vs flattening.
// ============================================================

/**
 * Make every property of T nullable and required.
 * ProseMirror stores every declared attr; explicit null matches that model.
 */
export type AttrNullable<T> = { [K in keyof T]-?: T[K] | null };

/**
 * Paragraph and heading attrs — mirrors ParagraphPropertiesOptionsBase.
 *
 * indent/spacing/border/run/frame are nested objects (matching office-open),
 * so one `indent` attr replaces 13 flattened indent attrs.
 */
export type ParagraphAttrs = AttrNullable<ParagraphPropertiesOptionsBase>;

/**
 * Text style mark attrs — mirrors RunStylePropertiesOptions.
 *
 * Omits properties handled by dedicated marks (bold, italic, strike,
 * doubleStrike, subScript, superScript). `size` is in POINTS.
 */
export type TextStyleAttrs = AttrNullable<
  Omit<
    RunStylePropertiesOptions,
    | "bold"
    | "boldComplexScript"
    | "italic"
    | "italicComplexScript"
    | "strike"
    | "doubleStrike"
    | "subScript"
    | "superScript"
  >
>;

/**
 * Link mark attrs.
 */
export interface LinkAttrs {
  href: string | null;
  target: string | null;
  rel: string | null;
  class: string | null;
  title: string | null;
}

/**
 * Table attrs — mirrors TableOptions (minus `rows`, which is structural).
 */
export type TableAttrs = AttrNullable<Omit<TableOptions, "rows">>;

/**
 * Table row attrs — mirrors TableRowPropertiesOptionsBase.
 * height is nested { value, rule } matching office-open.
 */
export type TableRowAttrs = AttrNullable<TableRowPropertiesOptionsBase>;

/**
 * Table cell / header attrs — mirrors TableCellOptions.
 *
 * colspan/rowspan/colwidth kept as Tiptap structural names (base extension
 * dependent); office-open columnSpan/rowSpan are mapped in renderDocx.
 */
export type TableCellAttrs = AttrNullable<
  Omit<TableCellOptions, "children" | "columnSpan" | "rowSpan">
> & {
  /** Horizontal span (Tiptap base name; maps to office-open columnSpan). */
  colspan: number;
  /** Vertical span (Tiptap base name; maps to office-open rowSpan). */
  rowspan: number;
  /** Column width in pixels per cell (Tiptap base name). */
  colwidth: number[] | null;
};

/**
 * Image attrs.
 * src/alt/title kept as Tiptap structural names.
 * width/height are pixel dimensions for editor display.
 */
export interface ImageAttrs {
  src: string;
  alt: string | null;
  title: string | null;
  width: number | null;
  height: number | null;
  rotation: number | null;
  floating: Floating | null;
  outline: NonNullable<ImageOptions["outline"]> | null;
  crop: NonNullable<ImageOptions["sourceRectangle"]> | null;
  display: string | null;
  // 0.9.7+ fidelity fields (office-open native; near-identity passthrough)
  nonVisualProperties: NonNullable<ImageOptions["nonVisualProperties"]> | null; // pic:cNvPr (id/name/descr)
  effectExtent: { l: number; t: number; r: number; b: number } | null; // wp:effectExtent EMUs
  graphicFrameLocks: NonNullable<ImageOptions["graphicFrameLocks"]> | null;
  blipEffects: NonNullable<ImageOptions["blipEffects"]> | null;
  useLocalDpi: boolean | null; // a14:useLocalDpi
  fill: NonNullable<ImageOptions["fill"]> | null;
  effects: NonNullable<ImageOptions["effects"]> | null;
  tile: NonNullable<ImageOptions["tile"]> | null;
  runPropertiesRawXml: string | null;
}

/**
 * Emoji node attrs.
 * `name` is the shortcode (base @tiptap/extension-emoji); `emoji` is the
 * resolved glyph cached for DOCX export (null when only the shortcode is known).
 */
export interface EmojiAttrs {
  name: string | null;
  emoji: string | null;
}

/**
 * Strike mark attrs.
 */
export interface StrikeAttrs {
  doubleStrike: boolean | null;
}

// ============================================================
// Layer 3: Tiptap JSON node types
//
// These describe the structure of JSONContent produced by our
// custom extensions. Useful for consumers who need typed access.
// ============================================================

// -- Inline content --

export interface TextNode {
  type: "text";
  text: string;
  marks?: Mark[];
}

export interface HardBreakNode {
  type: "hardBreak";
  marks?: Mark[];
}

// -- Mark types --

export type Mark =
  | { type: "bold" }
  | { type: "italic" }
  | { type: "underline" }
  | { type: "strike"; attrs?: StrikeAttrs }
  | { type: "code" }
  | { type: "subscript" }
  | { type: "superscript" }
  | { type: "highlight"; attrs?: { color?: string } }
  | { type: "textStyle"; attrs?: TextStyleAttrs }
  | { type: "link"; attrs?: LinkAttrs };

// -- Block nodes --

export interface ParagraphNode extends TiptapJSONContent {
  type: "paragraph";
  attrs?: ParagraphAttrs;
  content?: Array<TextNode | HardBreakNode | ImageNode | EmojiNode>;
}

export interface HeadingNode extends TiptapJSONContent {
  type: "heading";
  attrs: { level: 1 | 2 | 3 | 4 | 5 | 6 } & ParagraphAttrs;
  content?: Array<TextNode | HardBreakNode>;
}

export interface BlockquoteNode extends TiptapJSONContent {
  type: "blockquote";
  content?: Array<ParagraphNode>;
}

export interface CodeBlockNode extends TiptapJSONContent {
  type: "codeBlock";
  attrs?: { language?: string };
  content?: Array<TextNode>;
}

export interface HorizontalRuleNode extends TiptapJSONContent {
  type: "horizontalRule";
}

/**
 * Header/footer slots in Tiptap JSON — each slot is the JSONContent[] produced
 * by resolving that slot's SectionChild[] (paragraphs/tables/…). Mirrors
 * SectionOptions.headers/footers in the runtime model.
 */
export interface HeaderFooterSlots {
  default?: TiptapJSONContent[];
  first?: TiptapJSONContent[];
  even?: TiptapJSONContent[];
}

// -- List nodes --

export interface BulletListNode extends TiptapJSONContent {
  type: "bulletList";
  content?: Array<ListItemNode>;
}

export interface OrderedListNode extends TiptapJSONContent {
  type: "orderedList";
  attrs?: { start?: number; order?: number; type?: string | null };
  content?: Array<ListItemNode>;
}

export interface TaskListNode extends TiptapJSONContent {
  type: "taskList";
  content?: Array<TaskItemNode>;
}

export interface ListItemNode extends TiptapJSONContent {
  type: "listItem";
  content?: Array<ParagraphNode>;
}

export interface TaskItemNode extends TiptapJSONContent {
  type: "taskItem";
  attrs?: { checked?: boolean };
  content?: Array<ParagraphNode>;
}

// -- Table nodes --

export interface TableNode extends TiptapJSONContent {
  type: "table";
  attrs?: TableAttrs;
  content?: Array<TableRowNode>;
}

export interface TableRowNode extends TiptapJSONContent {
  type: "tableRow";
  attrs?: TableRowAttrs;
  content?: Array<TableCellNode | TableHeaderNode>;
}

export interface TableCellNode extends TiptapJSONContent {
  type: "tableCell";
  attrs?: TableCellAttrs;
  content?: Array<ParagraphNode>;
}

export interface TableHeaderNode extends TiptapJSONContent {
  type: "tableHeader";
  attrs?: TableCellAttrs;
  content?: Array<ParagraphNode>;
}

// -- Image node --

export interface ImageNode extends TiptapJSONContent {
  type: "image";
  attrs?: ImageAttrs;
}

/**
 * Drawing group (wpg) carried as an opaque blob — the full WpgGroupRunOptions
 * (pictures/shapes/nested groups + transform) round-trips verbatim. The editor
 * doesn't model the group interior; renderHTML paints the first picture only.
 */
export interface ImageGroupNode extends TiptapJSONContent {
  type: "imageGroup";
  attrs?: { wpgGroup: Record<string, unknown> | null };
}

// -- Emoji node (inline) --

export interface EmojiNode extends TiptapJSONContent {
  type: "emoji";
  attrs?: EmojiAttrs;
}

// -- Details node --

export interface DetailsNode extends TiptapJSONContent {
  type: "details";
  content?: Array<DetailsSummaryNode | DetailsContentNode>;
}

export interface DetailsSummaryNode extends TiptapJSONContent {
  type: "detailsSummary";
  content?: Array<TextNode | HardBreakNode>;
}

export interface DetailsContentNode extends TiptapJSONContent {
  type: "detailsContent";
  content?: Array<BlockNode>;
}

// -- Union types --

export type InlineContent = TextNode | HardBreakNode;
export type BlockNode =
  | ParagraphNode
  | HeadingNode
  | BlockquoteNode
  | CodeBlockNode
  | HorizontalRuleNode
  | BulletListNode
  | OrderedListNode
  | TaskListNode
  | TableNode
  | ImageNode
  | DetailsNode;
