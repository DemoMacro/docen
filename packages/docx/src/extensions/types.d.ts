import type {
  ParagraphChild,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  StylesOptions,
} from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

declare module "@tiptap/core" {
  interface NodeConfig<Options = any, Storage = any> {
    /**
     * DOCX serialization: Tiptap JSON node → DOCX opts, or null when the node
     * cannot be serialized (e.g. an image with no embedded data — DocxManager
     * then drops it). Each node extension defines this to convert its attrs to
     * DOCX properties.
     */
    renderDocx?: (node: JSONContent) => Record<string, unknown> | null;
    /**
     * DOCX deserialization: DOCX opts → Tiptap JSON attrs.
     * Each node extension defines this to convert DOCX properties back to attrs.
     */
    parseDocx?: (opts: Record<string, unknown>) => Record<string, unknown>;
    /**
     * Declarative block parse rule: recognize a SectionChild shape this node
     * owns and convert it to a JSONContent node. DocxManager walks every
     * extension's rule (in docxExtensions declaration order) during resolve, so
     * a custom block node plugs in by declaring this instead of forking
     * DocxManager. `match` false → skipped; `convert` null after a positive
     * match falls through to the next rule, then the paragraph/passthrough
     * fallbacks.
     */
    parseDocxBlock?: ParseBlockRule;
    /**
     * Declarative inline parse rule: recognize a ParagraphChild shape this
     * node/mark owns and convert it to inline content (a node, text[], or
     * null to fall through). DocxManager walks every extension's rule during
     * resolve, so a custom inline shape plugs in by declaring this instead of
     * forking DocxManager. Used by both Node extensions (image/tab/wpsShape/…)
     * and Mark containers (hyperlink/insertion/deletion, which yield text
     * nodes carrying the mark).
     */
    parseDocxInline?: ParseInlineRule;
    /**
     * Declarative paragraph parse rule: recognize a paragraph subtype this node
     * owns (heading/codeBlock) and convert it to a JSONContent node. DocxManager
     * walks every extension's rule during resolve before the plain-paragraph
     * fallback, so a custom paragraph subtype plugs in by declaring this instead
     * of forking DocxManager. thematicBreak (→ horizontalRule) stays in the
     * manager — it has no owning extension.
     */
    parseDocxParagraph?: ParseParagraphRule;
    /**
     * Declarative aggregator: claims consecutive paragraphs that belong to a
     * composite structure (a list tree, a blockquote) and rebuilds them as
     * nested JSONContent. DocxManager keeps a generic group-by-predicate loop;
     * the rule contributes the predicate (belongs) + the builder (build). A
     * custom composite plugs in by declaring this instead of forking DocxManager.
     */
    parseDocxAggregator?: ParseAggregatorRule;
  }

  interface ExtensionConfig<Options = any, Storage = any> {
    /** Declarative aggregator on a plain Extension (e.g. a list-tree rebuilder
     *  that spans bullet/ordered/task lists). See NodeConfig.parseDocxAggregator. */
    parseDocxAggregator?: ParseAggregatorRule;
  }

  interface MarkConfig<Options = any, Storage = any> {
    /**
     * DOCX serialization: mark attrs → run-level properties (merged into the
     * run's options). Each mark extension defines this to contribute rPr fields.
     */
    renderDocx?: (attrs: Record<string, unknown>) => Record<string, unknown>;
    /**
     * DOCX deserialization: RunOptions → mark attrs, or null when the run does
     * not carry this mark (DocxManager then skips emitting it). Each mark
     * extension defines this to extract its attrs from run properties.
     */
    parseDocx?: (opts: RunOptions) => Record<string, unknown> | null;
    /**
     * Declarative inline parse rule (see NodeConfig.parseDocxInline). A Mark
     * container such as hyperlink/insertion/deletion declares this to yield
     * text nodes carrying the mark.
     */
    parseDocxInline?: ParseInlineRule;
  }
}

/**
 * Per-resolve façade handed to parse rules: read-only views over the
 * DocumentOptions being resolved plus recursive entry points back into
 * DocxManager — a table cell, a TOC entry, and a details body are themselves
 * SectionChild[] block streams. Rule bodies must stay pure: getExtensionField
 * binds no `this`, so a rule cannot read extension options/storage and reaches
 * the manager only through this context.
 */
export interface ResolveContext {
  /** Walk a SectionChild[] block stream (a cell's children, a TOC's entries, a
   *  details body) — regroups consecutive numbering paragraphs into lists. */
  resolveBlockStream(children: SectionChild[]): JSONContent[];
  /** Resolve one SectionChild (a TOC entry, a details body block). */
  resolveBlock(child: SectionChild): JSONContent | null;
  /** Resolve a paragraph's inline content (handles the bare-string/{text}
   *  fallback office-open collapses a plain paragraph to). */
  resolveInlineContent(para: ParagraphOptions): JSONContent[];
  /** Walk a ParagraphChild[] inline stream (a hyperlink/track-change container's
   *  child runs) — the inline counterpart of resolveBlockStream. */
  resolveInlineChildren(children: (ParagraphChild | string)[]): JSONContent[];
  /** Resolve a paragraph (a wpsShape text-box body block). */
  resolveParagraph(para: ParagraphOptions | string): JSONContent;
  /** Reflective node attrs parse (table/tableRow/tableHeader/tableCell/…) —
   *  reuses the nodeParse registry so a block rule shares the attrs extraction
   *  the inline/compile paths use. */
  parseNodeAttrs(type: string, opts: Record<string, unknown>): Record<string, unknown>;
  /** Resolve run-level marks (bold/italic/…) for a RunOptions — used by
   *  code-block's resolveCodeBlock to recover inline marks on each run. */
  resolveMarks(opts: RunOptions): JSONContent["marks"];
  /** The document's styles.xml model. Table rules read tableStyles; paragraph
   *  rules (Phase 2) will read paragraphStyles. */
  readonly styles: StylesOptions | undefined;
  /** Numbering reference → level-0 format/start, for classifying list paragraphs
   *  (bullet vs ordered, start value). Read by the list aggregator. */
  readonly numberingLookup: Map<string, { format?: string; start?: number }> | undefined;
}

/** A declarative block parse rule. `match` identifies the SectionChild shape
 *  this node owns; `convert` builds the JSONContent node (null falls through to
 *  the next rule). */
export interface ParseBlockRule {
  match(child: SectionChild, ctx: ResolveContext): boolean;
  convert(child: SectionChild, ctx: ResolveContext): JSONContent | null;
}

/** A declarative inline parse rule. `match` identifies the ParagraphChild
 *  shape; `convert` builds inline content — a node, an array (hyperlink/
 *  track-change yield text[]), or null to fall through to the next rule. */
export interface ParseInlineRule {
  match(child: ParagraphChild, ctx: ResolveContext): boolean;
  convert(child: ParagraphChild, ctx: ResolveContext): JSONContent | JSONContent[] | null;
}

/** A declarative paragraph parse rule. `match` identifies the paragraph
 *  subtype; `convert` builds the JSONContent node (null falls through to the
 *  next rule, then the plain-paragraph fallback). */
export interface ParseParagraphRule {
  match(para: ParagraphOptions, ctx: ResolveContext): boolean;
  convert(para: ParagraphOptions, ctx: ResolveContext): JSONContent | null;
}

/** A declarative aggregator rule. `belongs` is the predicate grouping
 *  consecutive paragraphs into one composite; `build` turns the group into
 *  JSONContent[] (e.g. a nested list tree, a blockquote). The manager runs the
 *  generic group-by loop — the rule owns only the predicate + builder. */
export interface ParseAggregatorRule {
  belongs(para: ParagraphOptions, ctx: ResolveContext): boolean;
  build(group: ParagraphOptions[], ctx: ResolveContext): JSONContent[];
}
