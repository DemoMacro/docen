/**
 * The neutral Markdown IR: a lossless, format-agnostic shape of the Markdown
 * syntax itself (CommonMark + GFM tables). Format packages map this IR to and
 * from their own document models through {@link MarkdownMapper}; this package
 * knows no document model.
 */

/** Inline rich-text marks — the direct translation of Markdown emphasis
 *  syntax, shared across formats (DOCX runs, PPTX runs, XLSX rich text). */
export type MdMark =
  | { type: "bold" }
  | { type: "italic" }
  | { type: "strike" }
  | { type: "code" }
  | { type: "link"; href: string; title?: string };

export type MdInline =
  | { type: "text"; text: string; marks?: MdMark[] }
  | { type: "image"; src: string; alt?: string; title?: string }
  | { type: "hardBreak" };

export type MdAlignment = "left" | "center" | "right" | null;

export interface MdTableCell {
  /** The header row's cells (GFM delimiter row position). */
  header?: boolean;
  blocks: MdBlock[];
}

export interface MdTableRow {
  cells: MdTableCell[];
}

export interface MdListItem {
  blocks: MdBlock[];
  /** GFM task state (`- [ ]` / `- [x]`); absent for plain items. */
  checked?: boolean;
}

export type MdBlock =
  | { type: "heading"; depth: number; children: MdInline[] }
  | { type: "paragraph"; children: MdInline[] }
  | { type: "codeBlock"; text: string; lang?: string }
  | { type: "quote"; blocks: MdBlock[] }
  | { type: "list"; ordered: boolean; start?: number; items: MdListItem[] }
  | { type: "table"; align: MdAlignment[]; rows: MdTableRow[] }
  | { type: "thematicBreak" }
  | { type: "footnoteDefinition"; label: string; children: MdInline[] };

/**
 * Bridges the Markdown IR and one format's document model. Implemented by
 * each format package (`@docen/docx` ships the reference mapper; PPTX/XLSX
 * follow the same pattern); this package only defines the contract.
 */
export interface MarkdownMapper<T> {
  /** IR block → zero or more target nodes (a format may flatten or fold). */
  parseBlock(block: MdBlock): T[];
  /** IR inline run → target inline nodes. */
  parseInline(inlines: MdInline[]): T[];
  /** Target document → IR blocks (the reverse of parseBlock/parseInline). */
  toBlocks(doc: T): MdBlock[];
}
