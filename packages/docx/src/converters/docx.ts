import {
  generateDocument,
  generateDocumentStream,
  generateDocumentSync,
  parseDocument,
} from "@office-open/docx";
import type {
  DocumentOptions,
  SectionChild,
  SectionPropertiesOptions,
  ParagraphOptions,
  ParagraphChild,
  RunOptions,
  TableOptions,
  TableCellOptions,
  BorderOptions,
  LevelsOptions,
  OutputByType,
  OutputType,
  PackerOptions,
  StylesOptions,
  TableOfContentsOptions,
} from "@office-open/docx";
import { flattenExtensions, getExtensionField, getSchema } from "@tiptap/core";

import type { Extensions, JSONContent } from "../core";
import { docxExtensions } from "../core";
import { buildListLevels, isGeneratedListReference } from "../extensions/list-numbering";
import type { ParseBlockRule, ParseInlineRule, ResolveContext } from "../extensions/types";
import { prepareDocument, type PrepareStep } from "./prepare";
import { buildTextBlock } from "./styles";

export type { DocumentOptions };

// ── Helpers ──

/** True when two cell-margin sets (w:tcMar / w:tblCellMar) match on every side's
 *  size — used to detect a cell that merely echoes the table's default so its
 *  redundant tcMar can be dropped for a near-identity round-trip. */
function sameCellMargins(
  a: NonNullable<TableCellOptions["margins"]>,
  b: NonNullable<TableCellOptions["margins"]>,
): boolean {
  const sz = (
    m: NonNullable<TableCellOptions["margins"]>,
    k: "top" | "right" | "bottom" | "left",
  ) => m[k]?.size ?? null;
  return (["top", "right", "bottom", "left"] as const).every((k) => sz(a, k) === sz(b, k));
}

/** True when two single-side borders match on style/size/color — used to detect
 *  a cell side that merely echoes the table's insideHorizontal/insideVertical
 *  so it can be dropped for a near-identity round-trip (resolveTable pushes
 *  those table-level grid lines onto cell sides lacking their own tcBorder). */
function sameBorder(a: BorderOptions | undefined, b: BorderOptions | undefined): boolean {
  if (!a || !b) return false;
  return a.style === b.style && a.size === b.size && a.color === b.color;
}

/** Push one item or a spread of items onto target — block/inline parse results
 *  and section-child compiles legitimately return either shape. */
function pushAll<T>(target: T[], item: T | T[]): void {
  if (Array.isArray(item)) target.push(...item);
  else target.push(item);
}

/**
 * Core document properties (docProps/core.xml) carried on `doc.attrs.core` for
 * lossless round-trip. Mirrors @office-open/core's CorePropertiesOptions, which
 * @office-open/docx's DocumentOptions extends. Inlined here (not imported from
 * @office-open/core) to keep @docen/docx's dependency surface on @office-open/docx.
 */
interface DocxCoreProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  lastPrinted?: string;
  created?: string;
  modified?: string;
  revision?: number;
}

/** Keys round-tripped between DocumentOptions core properties and `doc.attrs.core`. */
const CORE_PROPERTY_KEYS: readonly (keyof DocxCoreProperties)[] = [
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastModifiedBy",
  "lastPrinted",
  "created",
  "modified",
  "revision",
];

/**
 * DocumentOptions keys that DocxManager reconstructs (sections/numbering) or
 * carries in dedicated attrs (styles/background/core). Excluded from the
 * documentExtras pass-through so they aren't duplicated. The satisfies guard
 * pins every listed key to a real DocumentOptions field — renaming or
 * inventing one fails the build (documentExtras stays an open set: its value
 * is auto-compat with new top-level office-open keys).
 */
const COMPILE_OWNED_KEYS = new Set<string>([
  "sections",
  "numbering",
  "styles",
  "background",
  ...CORE_PROPERTY_KEYS,
] satisfies (keyof DocumentOptions)[]);

/** Collect core properties present on DocumentOptions into a plain object. */
function extractCoreProperties(docOpts: DocumentOptions): DocxCoreProperties | null {
  // Double cast: DocumentOptions (no index signature) is not comparable
  // to Record<string, unknown> in a single assertion.
  const source = docOpts as unknown as Record<string, unknown>;
  const core: Record<string, unknown> = {};
  for (const key of CORE_PROPERTY_KEYS) {
    const value = source[key];
    if (value !== undefined && value !== null) core[key] = value;
  }
  return Object.keys(core).length > 0 ? (core as DocxCoreProperties) : null;
}

/**
 * Header/footer slot shapes for the two sides of the round-trip:
 * - {@link SectionHeaderFooterGroup} — persistence side (SectionChild[] per
 *   slot); matches SectionOptions.headers/footers.
 * - {@link HeaderFooterSlots} — runtime side (JSONContent[] per slot), produced
 *   by resolveSectionChildren and consumed by compileSectionChild.
 */
type SectionHeaderFooterGroup = {
  default?: SectionChild[];
  first?: SectionChild[];
  even?: SectionChild[];
};

interface HeaderFooterSlots {
  default?: JSONContent[];
  first?: JSONContent[];
  even?: JSONContent[];
}

// ── DocxManager ──

/**
 * Manages DOCX serialization (Tiptap JSON ↔ DocumentOptions).
 *
 * Each extension provides renderDocx/parseDocx for its own attrs mapping.
 * DocxManager handles tree walking, child assembly, and dispatching.
 */
export class DocxManager {
  // Numbering definitions registered during compile for editor-created lists
  // (a `docen-bullet`/`docen-ordered-*` reference no source definition covers).
  // Round-tripped references (`list_<numId>`) carry their definitions in the
  // source numbering and never land here.
  private numberingConfigs: { reference: string; levels: LevelsOptions[] }[] = [];
  // Generated list references seen on list paragraphs during this compile —
  // each gets a definition in numberingConfigs unless the source numbering
  // already defines it.
  private usedListReferences = new Set<string>();
  // Styles table (styles.xml) carried through resolve() so consumers can
  // resolve a NUMERIC pStyle whose NAME is a heading (style "2" → name
  // "heading 1") — office-open lifts only pStyle literals that ARE
  // HeadingLevels ("Heading1".."Title"); real DOCX files often use numeric
  // ids. Set per resolve(); compile never reads it.
  private resolveStyles: StylesOptions | undefined;
  // DOCX conversion hooks collected via reflection from the extension list.
  // Each mark/node extension declares renderDocx/parseDocx as config fields;
  // the constructor reads them with getExtensionField so user-supplied
  // extensions plug in without a fork. Container marks (link/insertion/
  // deletion) wrap runs and are handled directly in compile/resolve, not
  // through these maps. Node hooks cover attrs↔opts only; block-level children
  // assembly (table rows/cells, TOC entries) is owned by each
  // block extension's parseDocxBlock rule (blockRules below); inline content
  // and the compile-side tree walk stay in DocxManager.
  private markRender = new Map<
    string,
    (attrs: Record<string, unknown>) => Record<string, unknown>
  >();
  private markParse: Array<{
    name: string;
    parse: (opts: RunOptions) => Record<string, unknown> | null;
  }> = [];
  private nodeRender = new Map<string, (node: JSONContent) => Record<string, unknown> | null>();
  private nodeParse = new Map<string, (opts: object) => Record<string, unknown>>();
  // Declarative block parse rules collected via reflection (mirrors markParse).
  // Each block node extension declares parseDocxBlock; resolveSectionChild walks
  // them in docxExtensions order before the paragraph/passthrough fallbacks.
  private blockRules: Array<{ name: string; rule: ParseBlockRule }> = [];
  // Declarative inline parse rules collected via reflection (mirrors blockRules).
  // Each inline node/mark extension declares parseDocxInline; resolveParagraphChild
  // walks them in docxExtensions order before the run/sdt/passthrough fallbacks.
  private inlineRules: Array<{ name: string; rule: ParseInlineRule }> = [];
  // Per-resolve façade over the recursive resolve entry points + read-only
  // styles, handed to every block/inline rule. Built at the start of resolve().
  private resolveCtx: ResolveContext | undefined;

  constructor(extensions: Extensions = docxExtensions) {
    const schema = getSchema(extensions);
    for (const ext of flattenExtensions(extensions)) {
      const name = ext.name;
      if (!name) continue;
      if (schema.marks[name]) {
        const render = getExtensionField(ext, "renderDocx") as
          | ((attrs: Record<string, unknown>) => Record<string, unknown>)
          | undefined;
        const parse = getExtensionField(ext, "parseDocx") as
          | ((opts: RunOptions) => Record<string, unknown> | null)
          | undefined;
        if (render) this.markRender.set(name, render);
        if (parse) this.markParse.push({ name, parse });
      } else if (schema.nodes[name]) {
        const render = getExtensionField(ext, "renderDocx") as
          | ((node: JSONContent) => Record<string, unknown> | null)
          | undefined;
        const parse = getExtensionField(ext, "parseDocx") as
          | ((opts: object) => Record<string, unknown>)
          | undefined;
        if (render) this.nodeRender.set(name, render);
        if (parse) this.nodeParse.set(name, parse);
        const blockRule = getExtensionField(ext, "parseDocxBlock") as ParseBlockRule | undefined;
        if (blockRule) this.blockRules.push({ name, rule: blockRule });
      }
      // parseDocxInline is collected for both nodes and marks — inline shapes
      // include mark containers (hyperlink/insertion/deletion) that yield text[].
      const inlineRule = getExtensionField(ext, "parseDocxInline") as ParseInlineRule | undefined;
      if (inlineRule) this.inlineRules.push({ name, rule: inlineRule });
    }
  }

  /** Reflective node renderDocx lookup: the node's DOCX opts, or {} when the
   *  node type has no renderDocx hook (degrades to a plain paragraph). node.type
   *  is optional on JSONContent — an absent type simply misses the map. */
  private renderNodeOpts(node: JSONContent): Record<string, unknown> {
    return this.nodeRender.get(node.type ?? "")?.(node) ?? {};
  }

  compile(json: JSONContent): DocumentOptions {
    this.numberingConfigs = [];
    this.usedListReferences = new Set();

    // Split doc content into sections. A non-final section's sectPr attaches to
    // its LAST paragraph's pPr (OOXML) — that paragraph carries sectionProperties/
    // sectionHeaders/sectionFooters attrs and closes the section here. The
    // trailing blocks form the final section, whose sectPr rides on
    // doc.attrs.sectionProperties/sectionHeaders/sectionFooters (body-level).
    // No section-carrying paragraph → single section (backward compatible).
    const sections: DocumentOptions["sections"] = [];
    let currentChildren: SectionChild[] = [];
    if (json.content) {
      for (const node of json.content) {
        const child = this.compileSectionChild(node);
        if (child) pushAll(currentChildren, child);
        if (node.type === "paragraph") {
          const na = (node.attrs ?? {}) as Record<string, unknown>;
          if (na.sectionProperties != null) {
            sections.push(
              this.buildSection(
                currentChildren,
                na.sectionProperties as SectionPropertiesOptions | null,
                this.compileHeaderFooter((na.sectionHeaders ?? null) as HeaderFooterSlots | null),
                this.compileHeaderFooter((na.sectionFooters ?? null) as HeaderFooterSlots | null),
              ),
            );
            currentChildren = [];
          }
        }
      }
    }
    const docAttrs = json.attrs ?? {};
    sections.push(
      this.buildSection(
        currentChildren,
        (docAttrs.sectionProperties ?? null) as SectionPropertiesOptions | null,
        this.compileHeaderFooter((docAttrs.sectionHeaders ?? null) as HeaderFooterSlots | null),
        this.compileHeaderFooter((docAttrs.sectionFooters ?? null) as HeaderFooterSlots | null),
      ),
    );

    const styles = (docAttrs.styles ?? undefined) as DocumentOptions["styles"] | undefined;
    const core = (docAttrs.core ?? undefined) as DocxCoreProperties | undefined;
    const background = (docAttrs.background ?? undefined) as
      | DocumentOptions["background"]
      | undefined;
    const documentExtras = (docAttrs.documentExtras ?? undefined) as
      | Partial<DocumentOptions>
      | undefined;
    // Merge source numbering definitions (custom bullet/number markers) with
    // any regenerated ordered-list definitions; drop originals shadowed by a
    // regenerated reference to avoid duplicates.
    const origNumberingConfig =
      (
        docAttrs.numbering as
          | { abstractNumberings?: { reference: string; levels: LevelsOptions[] }[] }
          | undefined
      )?.abstractNumberings ?? [];
    // Register a generated reference's definition unless the source numbering
    // already carries one (an editor-created list re-toggling a round-tripped
    // reference must not shadow the original marker definition).
    for (const reference of this.usedListReferences) {
      if (!isGeneratedListReference(reference)) continue;
      if (origNumberingConfig.some((c) => c.reference === reference)) continue;
      const levels = buildListLevels(reference);
      if (levels) this.numberingConfigs.push({ reference, levels });
    }
    const regeneratedRefs = new Set(this.numberingConfigs.map((c) => c.reference));
    const numberingConfig = [
      ...origNumberingConfig.filter((c) => !regeneratedRefs.has(c.reference)),
      ...this.numberingConfigs,
    ];
    return {
      sections,
      ...(styles ? { styles } : {}),
      ...core,
      ...(background ? { background } : {}),
      ...documentExtras,
      ...(numberingConfig.length > 0 ? { numbering: { abstractNumberings: numberingConfig } } : {}),
    };
  }

  /** Assemble a SectionOptions from compiled children + optional layout/headers/footers. */
  private buildSection(
    children: SectionChild[],
    properties: SectionPropertiesOptions | null,
    headers: SectionHeaderFooterGroup | undefined,
    footers: SectionHeaderFooterGroup | undefined,
  ): DocumentOptions["sections"][number] {
    return {
      children,
      ...(properties ? { properties } : {}),
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
    };
  }

  /**
   * Compile resolved header/footer slots (JSONContent[] per slot) back into
   * SectionChild[] per slot. Returns undefined when no slot has content.
   */
  private compileHeaderFooter(
    slots: HeaderFooterSlots | null,
  ): SectionHeaderFooterGroup | undefined {
    if (!slots) return undefined;
    const group: SectionHeaderFooterGroup = {};
    for (const slot of ["default", "first", "even"] as const) {
      const json = slots[slot];
      if (!json?.length) continue;
      const children: SectionChild[] = [];
      for (const node of json) {
        const child = this.compileSectionChild(node);
        if (!child) continue;
        pushAll(children, child);
      }
      if (children.length > 0) group[slot] = children;
    }
    return Object.keys(group).length > 0 ? group : undefined;
  }

  /**
   * Resolve a section's header/footer group (SectionChild[] per slot) into
   * Tiptap JSON slots. Returns null when no slot has content.
   */
  private resolveHeaderFooter(
    group: SectionHeaderFooterGroup | undefined,
  ): HeaderFooterSlots | null {
    if (!group) return null;
    const slots: HeaderFooterSlots = {};
    for (const slot of ["default", "first", "even"] as const) {
      const children = group[slot];
      if (children?.length) {
        slots[slot] = this.resolveSectionChildren(children);
      }
    }
    return Object.keys(slots).length > 0 ? slots : null;
  }

  resolve(docOpts: DocumentOptions): JSONContent {
    this.resolveStyles = docOpts.styles ?? undefined;
    const sections = docOpts.sections ?? [];
    if (sections.length === 0) {
      return { type: "doc", content: [{ type: "paragraph" }] };
    }
    // Per-resolve façade handed to block parse rules. Arrow closures capture
    // `this`; `styles` is a read-only snapshot of the instance field set above
    // (stable for this resolve's lifetime).
    const styles = this.resolveStyles;
    this.resolveCtx = {
      resolveBlockStream: (children) => this.resolveSectionChildren(children),
      resolveBlock: (child) => this.resolveSectionChild(child),
      resolveInlineContent: (para) => this.resolveInlineContent(para),
      resolveInlineChildren: (children) => this.resolveParagraphChildren(children),
      resolveParagraph: (para) => this.resolveParagraph(para),
      parseNodeAttrs: (type, opts) => this.nodeParse.get(type)?.(opts) ?? {},
      resolveMarks: (opts) => this.resolveMarks(opts),
      styles,
    };

    // Resolve every section's children into blocks. A non-final section's
    // sectPr attaches to that section's LAST paragraph's pPr (OOXML) — stamp it
    // on that paragraph's attrs, not a standalone sectionBreak node. The final
    // section's sectPr rides on doc.attrs (it lives at <w:body>'s end).
    const content: JSONContent[] = [];
    const lastIndex = sections.length - 1;
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      const sectionContent = this.resolveSectionChildren(section.children ?? []);
      if (i < lastIndex) {
        const sectAttrs: Record<string, unknown> = {
          sectionProperties: section.properties ?? null,
          sectionHeaders: this.resolveHeaderFooter(section.headers),
          sectionFooters: this.resolveHeaderFooter(section.footers),
        };
        const last = sectionContent[sectionContent.length - 1];
        // Stamp sectPr on the section's last paragraph directly (a heading is
        // a paragraph too — same node) instead of appending a stray empty one.
        if (last?.type === "paragraph") {
          last.attrs = { ...last.attrs, ...sectAttrs };
        } else {
          // Section ends on a non-textblock (table/etc.) — Word still needs the
          // sectPr on a paragraph, so append an empty one to carry it.
          sectionContent.push({ type: "paragraph", attrs: sectAttrs });
        }
      }
      content.push(...sectionContent);
    }

    const doc: JSONContent = {
      type: "doc",
      content: content.length > 0 ? content : [{ type: "paragraph" }],
    };
    // Carry the styles library (styles.xml) and core properties (docProps/
    // core.xml) through the JSON for lossless round-trip. office-open regenerates
    // importedStyles/docDefaultsXml/latentStylesXml from `styles`, and writes
    // `core` back to docProps/core.xml.
    const attrs: Record<string, unknown> = {};
    if (docOpts.styles) attrs.styles = docOpts.styles;
    if (docOpts.background) attrs.background = docOpts.background;
    // Source numbering definitions (abstractNum) carried verbatim so list
    // markers round-trip; compile merges these with regenerated ordered defs.
    if (docOpts.numbering) attrs.numbering = docOpts.numbering;
    const core = extractCoreProperties(docOpts);
    if (core) attrs.core = core;
    const lastSection = sections[lastIndex];
    if (lastSection.properties) attrs.sectionProperties = lastSection.properties;
    const lastHeaders = this.resolveHeaderFooter(lastSection.headers);
    if (lastHeaders) attrs.sectionHeaders = lastHeaders;
    const lastFooters = this.resolveHeaderFooter(lastSection.footers);
    if (lastFooters) attrs.sectionFooters = lastFooters;
    // Pass through document-level fields DocxManager doesn't reconstruct
    // (settings.xml flags like displayBackgroundShape, zoom, fonts, footnotes,
    // customProperties, …). Word needs displayBackgroundShape to render the
    // <w:background> element, so losing it makes the page background invisible.
    const documentExtras: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(docOpts as unknown as Record<string, unknown>)) {
      if (COMPILE_OWNED_KEYS.has(k)) continue;
      if (v === undefined || v === null) continue;
      documentExtras[k] = v;
    }
    if (Object.keys(documentExtras).length > 0) attrs.documentExtras = documentExtras;
    if (Object.keys(attrs).length > 0) doc.attrs = attrs;
    return doc;
  }

  // ── Compile: Tiptap JSON → DocumentOptions ──

  private compileSectionChild(node: JSONContent): SectionChild | SectionChild[] | null {
    switch (node.type) {
      case "paragraph":
        return { paragraph: this.compileParagraphNode(node) };
      case "table":
        return { table: this.compileTableNode(node) };
      case "image": {
        const imageRun = this.nodeRender.get(node.type)?.(node) ?? null;
        if (!imageRun) return null;
        return { paragraph: { children: [imageRun] } };
      }
      case "tocField": {
        const options = (node.attrs?.options as TableOfContentsOptions | undefined) ?? {};
        const entries: SectionChild[] = [];
        for (const child of node.content ?? []) {
          const compiled = this.compileSectionChild(child);
          if (!compiled) continue;
          pushAll(entries, compiled);
        }
        return { toc: { ...options, entries } };
      }
      case "passthrough": {
        // Opaque SectionChild (rawXml/bookmark/toc/textbox/…) round-tripped verbatim.
        const data = (node.attrs?.data as string) ?? "{}";
        try {
          return JSON.parse(data) as SectionChild;
        } catch {
          return null;
        }
      }
      default:
        return null;
    }
  }

  private compileParagraphNode(node: JSONContent): ParagraphOptions | string {
    // A list paragraph references its numbering definition by attr — collect
    // the reference so compile registers a generated definition for it.
    const numRef = (node.attrs?.numbering as { reference?: string } | null | undefined)?.reference;
    if (typeof numRef === "string" && numRef) this.usedListReferences.add(numRef);
    const opts = this.renderNodeOpts(node);
    const childList = this.compileInlineContent(node.content);
    if (childList.length > 0) opts.children = childList;
    return this.simplifyParagraph(opts);
  }

  /** Simple text optimization: merge plain runs into the text field (a
   *  children-less paragraph with no other properties collapses to a plain
   *  string — SectionChild.paragraph accepts both). */
  private simplifyParagraph(opts: Record<string, unknown>): ParagraphOptions | string {
    const children = opts.children as ParagraphChild[] | undefined;
    if (!children || children.length === 0) return opts as ParagraphOptions;

    const allSimpleText = children.every(
      (c) =>
        typeof c === "object" && c !== null && "text" in c && Object.keys(c as object).length === 1,
    );
    if (allSimpleText) {
      const combined = children.map((c) => (c as { text: string }).text).join("");
      delete opts.children;
      if (combined && Object.keys(opts).length === 0) return combined;
      if (combined) opts.text = combined;
    }

    return opts as ParagraphOptions;
  }

  private compileTableNode(node: JSONContent): TableOptions {
    const opts = this.renderNodeOpts(node);
    const colCount = this.getTableColumnCount(node);
    // Table-level tblCellMar default — passed to compileTableCellNode so it can
    // drop a cell tcMar that merely echoes it (see resolveTable's push-down).
    const tableCellMargins = (opts.cellMargin ?? opts.margins ?? null) as NonNullable<
      TableCellOptions["margins"]
    > | null;
    // Table-level insideH/V — passed to compileTableCellNode so it can drop a
    // cell side that merely echoes them (resolveTable pushed them onto cells).
    const tableBorders = (opts.borders ?? null) as {
      insideHorizontal?: BorderOptions;
      insideVertical?: BorderOptions;
    } | null;
    const insideH = tableBorders?.insideHorizontal ?? null;
    const insideV = tableBorders?.insideVertical ?? null;
    // Compute tblGrid column widths up front so each cell's tcW can be derived
    // from the grid columns it spans. A cell without its own width attr would
    // otherwise collapse to the library default and autofit tables with short
    // content render as a sliver.
    const { columnWidths, tableWidth } =
      colCount > 0
        ? this.computeColumnWidths(node, colCount)
        : { columnWidths: [] as number[], tableWidth: { size: 0, type: "twips" } };
    const rows: Record<string, unknown>[] = [];

    // Cell attrs already mirror TableCellOptions (columnSpan/verticalMerge/
    // width included), so rows/cells pass straight through — no vMerge
    // interleave or rowspan rebuild. A cell's tcW is still derived from the
    // tblGrid columns it spans when it carries no width of its own.
    for (const rowNode of node.content ?? []) {
      if (rowNode.type !== "tableRow") continue;

      const rowOpts = this.nodeRender.get(rowNode.type)?.(rowNode) ?? {};
      const compiledCells: Record<string, unknown>[] = [];
      let colIdx = 0;
      for (const cellNode of rowNode.content ?? []) {
        if (cellNode.type !== "tableCell") continue;
        compiledCells.push(
          this.compileTableCellNode(
            cellNode,
            tableCellMargins,
            insideH,
            insideV,
            columnWidths,
            colIdx,
          ),
        );
        colIdx += (cellNode.attrs?.columnSpan as number) ?? 1;
      }
      rowOpts.cells = compiledCells;
      rows.push(rowOpts);
    }

    opts.rows = rows;

    if (colCount > 0) {
      opts.columnWidths = columnWidths;
      if (!opts.width) opts.width = tableWidth;
      if (!opts.layout) opts.layout = "autofit";
    }

    // Double cast: TableOptions requires `rows`, which opts gains only at
    // runtime from the loop above.
    return opts as unknown as TableOptions;
  }

  /** Count grid columns from the first table row (summing columnSpan). */
  private getTableColumnCount(tableNode: JSONContent): number {
    for (const rowNode of tableNode.content ?? []) {
      if (rowNode.type !== "tableRow") continue;
      let count = 0;
      for (const cell of rowNode.content ?? []) {
        if (cell.type === "tableCell") {
          count += (cell.attrs?.columnSpan as number) ?? 1;
        }
      }
      return count;
    }
    return 0;
  }

  /** Sum the tblGrid widths covered by a cell starting at `start` and spanning
   *  `span` grid columns — the OOXML tcW for that cell. */
  private sumGridSpan(columnWidths: number[], start: number, span: number): number {
    let sum = 0;
    for (let i = start; i < start + span && i < columnWidths.length; i++) {
      sum += columnWidths[i] ?? 0;
    }
    return sum;
  }

  /**
   * Compute tblGrid column widths (twips). The table's columnWidths attr
   * (tblGrid from a DOCX round-trip or insert-table) carries exact twips;
   * without it columns split the A4 content width evenly and the table
   * takes 100% width.
   */
  private computeColumnWidths(
    tableNode: JSONContent,
    colCount: number,
  ): { columnWidths: number[]; tableWidth: { size: number; type: string } } {
    const DEFAULT_CONTENT_TWIPS = 9026; // A4 with 1-inch margins (~15.9cm)

    const tableColWidths = tableNode.attrs?.columnWidths as number[] | null | undefined;
    if (tableColWidths && tableColWidths.length > 0) {
      const filled = tableColWidths.slice(0, colCount);
      while (filled.length < colCount) {
        filled.push(filled[filled.length - 1] ?? Math.floor(DEFAULT_CONTENT_TWIPS / colCount));
      }
      const total = filled.reduce((a, b) => a + b, 0);
      return { columnWidths: filled, tableWidth: { size: total, type: "twips" } };
    }

    // No explicit widths → equal distribution, percentage table width
    const equal = Math.floor(DEFAULT_CONTENT_TWIPS / colCount);
    return {
      columnWidths: Array(colCount).fill(equal),
      tableWidth: { size: 5000, type: "pct" }, // 100%
    };
  }

  private compileTableCellNode(
    cellNode: JSONContent,
    tableMargins?: NonNullable<TableCellOptions["margins"]> | null,
    insideH?: BorderOptions | null,
    insideV?: BorderOptions | null,
    columnWidths?: number[] | null,
    colIdx?: number,
  ): Record<string, unknown> {
    const cellOpts = this.renderNodeOpts(cellNode);

    // Restore the table-level form: resolveTable pushed the table's tblCellMar
    // default onto cells without their own tcMar. A cell whose margins equal
    // that default is the inherited one — drop its cell-level tcMar so the
    // regenerated docx keeps the compact table-level tblCellMar (near-identity
    // round-trip) instead of duplicating tcMar on every cell.
    if (
      tableMargins &&
      cellOpts.margins &&
      sameCellMargins(cellOpts.margins as NonNullable<TableCellOptions["margins"]>, tableMargins)
    ) {
      delete cellOpts.margins;
    }

    // Restore the table-level form: resolveTable pushed the table's insideH/V
    // onto cell sides lacking their own tcBorder. A side equal to the table's
    // insideH/V is the inherited one — drop it so the regenerated docx keeps
    // tblBorders.insideH/V instead of duplicating as tcBorders on every cell.
    if ((insideH || insideV) && cellOpts.borders) {
      const b = cellOpts.borders as Record<string, BorderOptions | undefined>;
      if (insideH && sameBorder(b.top, insideH)) delete b.top;
      if (insideH && sameBorder(b.bottom, insideH)) delete b.bottom;
      if (insideV && sameBorder(b.left, insideV)) delete b.left;
      if (insideV && sameBorder(b.right, insideV)) delete b.right;
      if (Object.keys(b).length === 0) delete cellOpts.borders;
    }

    // A cell may contain ANY block (nested table/list/…), not just
    // paragraphs. Route each child through the shared block compiler so a
    // nested list or table survives the round-trip.
    const cellChildren: SectionChild[] = [];
    for (const childNode of cellNode.content ?? []) {
      const compiled = this.compileSectionChild(childNode);
      if (compiled == null) continue;
      pushAll(cellChildren, compiled);
    }
    if (cellChildren.length > 0) cellOpts.children = cellChildren;

    // tcW = sum of the tblGrid columns the cell spans. tblGrid carries the
    // authoritative column widths — recomputing from it keeps autofit tables
    // with short content from collapsing to a sliver when the parsed tcW
    // disagrees.
    if (columnWidths && columnWidths.length > 0 && colIdx != null) {
      const span = (cellNode.attrs?.columnSpan as number) ?? 1;
      const tw = this.sumGridSpan(columnWidths, colIdx, span);
      if (tw > 0) cellOpts.width = { size: tw, type: "twips" };
    }

    return cellOpts;
  }

  // ── Inline content ──

  private compileInlineContent(content?: JSONContent[]): ParagraphChild[] {
    if (!content) return [];
    const children: ParagraphChild[] = [];

    for (const node of content) {
      switch (node.type) {
        case "text":
          this.compileTextNode(node, children);
          break;
        case "hardBreak":
          children.push({ break: 1 });
          break;
        case "pageBreak":
          children.push({ pageBreak: true });
          break;
        case "columnBreak":
          children.push({ columnBreak: true });
          break;
        case "tab":
          // office-open emits <w:tab/> from a top-level tab:true run, but its
          // ParagraphChild union doesn't list the top-level shape — double
          // assertion bridges the .d.ts gap.
          children.push({ tab: true } as Record<string, unknown> as ParagraphChild);
          break;
        case "inlinePassthrough": {
          // Opaque inline ParagraphChild (bookmark/range markers, …) carried
          // verbatim — reverse of resolveParagraphChild's fallback.
          const data = (node.attrs?.data as string) ?? "{}";
          try {
            const parsed = JSON.parse(data) as ParagraphChild;
            if (parsed) children.push(parsed);
          } catch {
            /* malformed JSON — drop */
          }
          break;
        }
        case "image": {
          const imageRun = this.nodeRender.get(node.type)?.(node) ?? null;
          if (imageRun) children.push(imageRun);
          break;
        }
        case "wpgGroup": {
          // Same .d.ts gap as tab:true above — office-open consumes a wpgGroup
          // run but lists the branch only in TrackChangeChild, not here.
          const wpgGroup = node.attrs?.wpgGroup;
          if (wpgGroup) children.push({ wpgGroup } as unknown as ParagraphChild);
          break;
        }
        case "wpsShape": {
          // Editable text body: compile each content paragraph back to a
          // ParagraphOptions and reattach under wpsShape.children. Mirrors the
          // tocField compile (compileSectionChild → unwrap .paragraph).
          // Same .d.ts gap as tab:true above for the ParagraphChild union.
          const geometry = (node.attrs?.wpsShape ?? {}) as Record<string, unknown>;
          const body: (ParagraphOptions | string)[] = [];
          for (const child of node.content ?? []) {
            const compiled = this.compileSectionChild(child);
            if (!compiled) continue;
            const items = Array.isArray(compiled) ? compiled : [compiled];
            for (const it of items) {
              if (it && typeof it === "object" && "paragraph" in (it as object)) {
                body.push((it as { paragraph: ParagraphOptions | string }).paragraph);
              }
            }
          }
          children.push({
            wpsShape: { ...geometry, children: body },
          } as unknown as ParagraphChild);
          break;
        }
      }
    }

    return children;
  }

  private compileTextNode(node: JSONContent, children: ParagraphChild[]): void {
    const text = node.text ?? "";
    if (!text) return;

    // Split on "\n": OOXML ignores a literal "\n" inside <w:t>, so each newline
    // becomes a {break:1} run. Shared by paragraphs and code blocks — codeBlock
    // needs no special-case newline handling.
    const segments = text.split("\n");
    for (let i = 0; i < segments.length; i++) {
      if (i > 0) children.push({ break: 1 });
      if (segments[i]) this.compileTextRun(segments[i], node.marks, children);
    }
  }

  /** Emit a single run for `text` with all inline marks applied. Container
   *  marks (hyperlink/insertion/deletion) wrap the run AFTER the rPr overlay
   *  marks are collected — a link whose run carries a character style (e.g.
   *  the stamped "Hyperlink") must keep it inside the container, so the
   *  container legs can no longer short-circuit the overlay pass. */
  private compileTextRun(
    text: string,
    marks: JSONContent["marks"],
    children: ParagraphChild[],
  ): void {
    const runOpts: Record<string, unknown> = { text };
    let linkMark: NonNullable<JSONContent["marks"]>[number] | undefined;
    let trackMark: NonNullable<JSONContent["marks"]>[number] | undefined;

    for (const mark of marks ?? []) {
      if (mark.type === "link") {
        if (mark.attrs?.href) linkMark = mark;
        continue;
      }
      if (mark.type === "insertion" || mark.type === "deletion") {
        trackMark = mark;
        continue;
      }
      // rPr overlay marks — each extension's renderDocx contributes run props.
      const render = this.markRender.get(mark.type);
      if (render) Object.assign(runOpts, render((mark.attrs ?? {}) as Record<string, unknown>));
    }

    if (linkMark) {
      const href = linkMark.attrs?.href as string;
      const { text: _, ...runWithoutText } = runOpts;
      const linkChildren: (RunOptions | string)[] = [];
      if (text) linkChildren.push({ ...runWithoutText, text } as RunOptions);
      children.push({
        hyperlink: {
          url: href.startsWith("#") ? undefined : href,
          anchor: href.startsWith("#") ? href.slice(1) : undefined,
          // resolveHyperlink parks w:tooltip on the mark's title attr.
          tooltip: (linkMark.attrs?.title as string | null | undefined) ?? undefined,
          children: linkChildren,
        },
      });
      return;
    }
    if (trackMark) {
      // Wrap the run back into a w:ins/w:del container — the reverse of
      // resolveTrackedChange. compileTrackedChangeRun returns the typed
      // ParagraphChild branch, so no cast is needed here. office-open's
      // stringifyDeletedRun emits <w:delText> for deletion children.
      const kind = trackMark.type;
      if (kind === "insertion" || kind === "deletion") {
        children.push(this.compileTrackedChangeRun(kind, trackMark.attrs, text, runOpts));
        return;
      }
    }

    children.push(runOpts as RunOptions);
  }

  /**
   * Wrap a run back into a w:ins/w:del container — the reverse of
   * resolveTrackedChange. A literal-key ternary (`{insertion: body}` /
   * `{deletion: body}`) lets TS narrow to the ParagraphChild branch without a
   * cast, and `typeof` guards read `attrs` type-safely (no `as number/string`).
   * stringifyDeletedRun emits `<w:delText>` automatically for deletion children.
   */
  private compileTrackedChangeRun(
    type: "insertion" | "deletion",
    attrs: Record<string, unknown> | undefined,
    text: string,
    runOpts: Record<string, unknown>,
  ): ParagraphChild {
    const { text: _, ...runWithoutText } = runOpts;
    const trackChildren: (RunOptions | string)[] = [];
    if (text) trackChildren.push({ ...runWithoutText, text } as RunOptions);
    const id = typeof attrs?.id === "number" ? attrs.id : 0;
    const author = typeof attrs?.author === "string" ? attrs.author : "";
    const date = typeof attrs?.date === "string" ? attrs.date : "";
    const body = { id, author, date, children: trackChildren };
    return type === "insertion" ? { insertion: body } : { deletion: body };
  }

  // ── Resolve: DocumentOptions → Tiptap JSON ──

  private resolveSectionChild(child: SectionChild): JSONContent | null {
    // Declarative block dispatch: each block extension's parseDocxBlock rule
    // (collected in docxExtensions order) gets a chance to recognize the shape.
    // table/toc own their shapes here; a non-matching or null-converting
    // rule falls through. The shapes are mutually exclusive (different
    // SectionChild keys), so order among them is irrelevant in practice.
    const ctx = this.resolveCtx!;
    for (const { rule } of this.blockRules) {
      if (rule.match(child, ctx)) {
        const node = rule.convert(child, ctx);
        if (node) return node;
      }
    }
    // paragraph is not a block rule — the paragraph subtype resolver and the
    // plain-paragraph fallback own it.
    if ("paragraph" in child) {
      return this.resolveParagraph(child.paragraph);
    }
    // rawXml (incl. aggregated TOC field), generic SDT, bookmarkStart/End,
    // textbox, altChunk, subDoc, customXml — no native Tiptap node. Carry the
    // SectionChild verbatim so the round-trip is byte-faithful.
    return this.resolvePassthrough(child);
  }

  /** Wrap an opaque SectionChild in a passthrough atom (attrs.data = JSON). */
  private resolvePassthrough(child: SectionChild): JSONContent {
    return { type: "passthrough", attrs: { data: JSON.stringify(child) } };
  }

  private resolveParagraph(opts: string | ParagraphOptions): JSONContent {
    const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;

    // Plain paragraph: reflective attrs parse + inline content. A paragraph
    // subtype (heading/code) is just attrs on this node.
    return buildTextBlock("paragraph", resolved, this.resolveCtx!);
  }

  /**
   * Walk a SectionChild[] block stream — a section's body, a header/footer
   * slot, or a table cell's children (a cell is just another block stream).
   */
  private resolveSectionChildren(children: SectionChild[]): JSONContent[] {
    const content: JSONContent[] = [];
    for (const child of children) {
      const node = this.resolveSectionChild(child);
      if (!node) continue;
      pushAll(content, node);
    }
    return content;
  }

  // ── Inline content resolution ──

  /**
   * Resolve a paragraph's inline content. @office-open collapses a plain-text
   * paragraph (a single run with no properties) to a bare string or a `{ text }`
   * object with no `children` — recover that text here so it round-trips back to
   * a text node instead of being dropped.
   */
  private resolveInlineContent(opts: ParagraphOptions): JSONContent[] {
    const content = this.resolveParagraphChildren(opts.children);
    if (content.length === 0 && opts.text) {
      const marks = this.resolveMarks(opts as unknown as RunOptions);
      const node: JSONContent = { type: "text", text: opts.text };
      if (marks) node.marks = marks;
      return [node];
    }
    return content;
  }

  private resolveParagraphChildren(children?: (ParagraphChild | string)[]): JSONContent[] {
    if (!children || children.length === 0) return [];

    const nodes: JSONContent[] = [];
    for (const child of children) {
      if (typeof child === "string") {
        if (child) nodes.push({ type: "text", text: child });
        continue;
      }
      if (typeof child === "object" && child !== null) {
        const resolved = this.resolveParagraphChild(child);
        if (resolved) pushAll(nodes, resolved);
      }
    }
    return nodes;
  }

  private resolveParagraphChild(child: ParagraphChild): JSONContent | JSONContent[] | null {
    // Declarative inline dispatch: each inline node/mark extension's
    // parseDocxInline rule (collected in docxExtensions order) gets a chance to
    // recognize the shape. tab/image/wpgGroup/wpsShape/hyperlink/
    // insertion/deletion/pageBreak/columnBreak own their shapes here; a
    // non-matching or null-converting rule falls through. The shapes are mutually
    // exclusive (different ParagraphChild keys), so order among them is
    // irrelevant in practice.
    const ctx = this.resolveCtx!;
    for (const { rule } of this.inlineRules) {
      if (rule.match(child, ctx)) {
        const node = rule.convert(child, ctx);
        if (node) return node;
      }
    }
    // run catch-all: a plain run (text/children/break). Left in the manager — it
    // is the fallback every non-owned shape reaches, not an owned shape itself.
    if ("text" in child || "children" in child || "break" in child) {
      return this.resolveRun(child as RunOptions);
    }
    // Any remaining inline shape (an inline SDT, bookmark/range
    // markers, proofErr, …) carries verbatim via inlinePassthrough so the
    // round-trip stays byte-faithful — mirrors block resolvePassthrough, which
    // keeps every unrecognized SectionChild instead of dropping it. An
    // inline SDT used to be dropped here; carrying it restores the symmetry.
    return { type: "inlinePassthrough", attrs: { data: JSON.stringify(child) } };
  }

  private resolveRun(opts: RunOptions): JSONContent | JSONContent[] | null {
    // Pure break (no text/children) → hardBreak node
    if (opts.break && opts.text === undefined && !opts.children) {
      return { type: "hardBreak" };
    }
    const text = opts.text;
    if (text === undefined && !opts.children) return null;

    // office-open 0.10.11+ may nest a run's inline elements under run.children
    // (empty run elements like tab/date/lastRenderedPageBreak, or mixed breaks).
    // Walk the children so pageBreak/columnBreak/hardBreak atoms are still
    // yielded; text fragments join into one text node carrying the run's marks.
    if (opts.children) {
      const marks = this.resolveMarks(opts);
      const nodes: JSONContent[] = [];
      let parts: string[] = [];
      const flushText = () => {
        if (parts.length > 0) {
          const node: JSONContent = { type: "text", text: parts.join("") };
          if (marks) node.marks = marks;
          nodes.push(node);
          parts = [];
        }
      };
      for (const c of opts.children) {
        if (typeof c === "string") {
          parts.push(c);
        } else if (c && typeof c === "object") {
          // Reflective: reuse the inlineRules dispatch (same as top-level
          // ParagraphChild) so tab/pageBreak/columnBreak — and any custom
          // inline atom — are recognized here too. This replaces a parallel
          // if-in chain that duplicated those rules and dropped shapes that
          // lacked a hardcoded branch (e.g. a custom inline node nested in a
          // run was silently lost).
          const ctx = this.resolveCtx!;
          let handled = false;
          for (const { rule } of this.inlineRules) {
            if (rule.match(c as ParagraphChild, ctx)) {
              const node = rule.convert(c as ParagraphChild, ctx);
              if (node) {
                flushText();
                pushAll(nodes, node);
                handled = true;
                break;
              }
            }
          }
          if (handled) continue;
          // hardBreak: a bare `<w:br/>` inside run.children — no owning inline
          // rule (HardBreak declares no parseDocxInline), stays a manager case.
          if ("break" in c) {
            flushText();
            nodes.push({ type: "hardBreak" });
          }
          // {lastRenderedPageBreak} is a Word render hint — drop (office-open
          // does not emit it on output). noBreakHyphen/date fields/separator/pgNum
          // are unsupported inline elements, dropped for now.
        }
      }
      flushText();
      if (text !== undefined) {
        // A run may carry opts.text alongside children (rare); fold it in.
        const node: JSONContent = { type: "text", text, marks };
        nodes.push(node);
      }
      if (nodes.length === 0) return null;
      return nodes.length === 1 ? nodes[0] : nodes;
    }

    return { type: "text", text: text ?? "", marks: this.resolveMarks(opts) };
  }

  private resolveMarks(opts: RunOptions): JSONContent["marks"] {
    const marks: NonNullable<JSONContent["marks"]> = [];

    // Each mark extension's parseDocx returns its attrs, or null when the run
    // does not carry the mark. The code/textStyle coupling (rStyle "CodeChar"
    // belongs to `code`; Consolas font and the styleId are skipped inside
    // textStyle.parseDocx) is handled within those extensions.
    for (const { name, parse } of this.markParse) {
      const attrs = parse(opts);
      if (attrs === null) continue;
      marks.push(Object.keys(attrs).length ? { type: name, attrs } : { type: name });
    }

    return marks.length > 0 ? marks : undefined;
  }
}

// ── Standalone functions (backward compat) ──

const defaultManager = new DocxManager(docxExtensions);

/** Resolve a DocxManager for a conversion call: the shared default singleton
 *  when no extensions are given, or a fresh instance bound to custom extensions
 *  so user-supplied marks/nodes plug into compile/resolve without a fork. */
function getDocxManager(extensions?: Extensions): DocxManager {
  return extensions ? new DocxManager(extensions) : defaultManager;
}

/**
 * Parse a DOCX file into Tiptap JSON (runtime model).
 *
 * Combines @office-open/docx's `parseDocument` (DOCX binary → DocumentOptions)
 * with `DocxManager.resolve` (DocumentOptions → Tiptap JSON).
 */
export function parseDOCX(
  data: Parameters<typeof parseDocument>[0],
  extensions?: Extensions,
): JSONContent {
  return getDocxManager(extensions).resolve(parseDocument(data));
}

/**
 * Options for {@link generateDOCX} / {@link generateDOCXStream}.
 */
export interface DocxGenerateOptions<T extends OutputType = "nodebuffer"> {
  /**
   * Pre-compilation steps run on the JSON in place (default: `prepareImages()`).
   * - `true` / `undefined`: default image pre-fetch (http(s) → embedded data URL)
   * - `false`: skip preparation
   * - `PrepareStep[]`: custom steps
   *
   * Required for http image URLs — image `renderDocx` drops images without
   * embedded data (see extensions/image.ts). Mutates the JSON, like `prepareDocument`.
   */
  prepare?: boolean | PrepareStep[];
  /** Packer options; `type` controls the output format (default `"nodebuffer"` → Buffer). */
  packer?: PackerOptions<T>;
  /**
   * Document-level options injected into the compiled `DocumentOptions` — core
   * properties (`title`/`creator`/`description`/…), `styles`, `background`,
   * `features`, `fonts`, etc. Excludes `sections` (always compiled from the
   * JSON) and `numbering` (collected from ordered-list nodes).
   *
   * `styles` here takes precedence over any `styles` carried on
   * `json.attrs.styles` (e.g. from a prior `parseDOCX`).
   */
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">;
  /**
   * Extension list used to build the conversion registry (default:
   * `docxExtensions`). Pass `[...docxExtensions, MyMark]` to plug a custom
   * mark/node into compile/resolve via its renderDocx/parseDocx hooks.
   */
  extensions?: Extensions;
}

/**
 * Merge {@link DocxGenerateOptions.document} into the `DocumentOptions`
 * compiled from Tiptap JSON.
 *
 * - `sections`/`numbering`: compile-owned (excluded from `document`, never
 *   overridden).
 * - `styles`: option wins over `json.attrs.styles`.
 * - Everything else (core properties, background, features, …): injected.
 */
function applyDocumentOptions(
  base: DocumentOptions,
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">,
): DocumentOptions {
  if (!document) return base;
  return { ...base, ...document };
}

/**
 * Generate a DOCX file from Tiptap JSON (runtime model), asynchronously.
 *
 * Pipeline: `prepareDocument` (default: fetch http images, in place) →
 * `DocxManager.compile` → @office-open/docx's `generateDocument`. `packer.type`
 * controls the output format (default: `"nodebuffer"` → Buffer). Non-blocking
 * (fflate Web Workers). With the default `prepare`, the input `json` is mutated
 * in place (http image URLs become embedded data URLs).
 */
export async function generateDOCX<T extends OutputType = "nodebuffer">(
  json: JSONContent,
  options?: DocxGenerateOptions<T>,
): Promise<OutputByType[T]> {
  const { prepare = true, packer, document, extensions } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocument(
    applyDocumentOptions(compileDocument(json, extensions), document),
    packer,
  );
}

/**
 * Generate a DOCX file synchronously — fastest throughput, blocks the event loop.
 *
 * Pipeline: `DocxManager.compile` → `generateDocumentSync`. Does **not** run
 * `prepareDocument` (it is async); call `await prepareDocument(json)` first
 * when http images need embedding. `options.document` is still applied.
 */
export function generateDOCXSync<T extends OutputType = "nodebuffer">(
  json: JSONContent,
  options?: DocxGenerateOptions<T>,
): OutputByType[T] {
  const { packer, document, extensions } = options ?? {};
  return generateDocumentSync(
    applyDocumentOptions(compileDocument(json, extensions), document),
    packer,
  );
}

/**
 * Generate a DOCX file as a `ReadableStream<Uint8Array>` — for large documents
 * or streaming HTTP responses.
 *
 * Pipeline: `prepareDocument` (default: fetch http images, in place) →
 * `DocxManager.compile` → `generateDocumentStream`. Async due to preparation.
 */
export async function generateDOCXStream(
  json: JSONContent,
  options?: DocxGenerateOptions,
): Promise<ReadableStream<Uint8Array>> {
  const { prepare = true, packer, document, extensions } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocumentStream(
    applyDocumentOptions(compileDocument(json, extensions), document),
    packer,
  );
}

/**
 * Convert DocumentOptions (persistence model) to Tiptap JSON (runtime model).
 */
export function resolveDocument(docOpts: DocumentOptions, extensions?: Extensions): JSONContent {
  return getDocxManager(extensions).resolve(docOpts);
}

/**
 * Convert Tiptap JSON (runtime model) to DocumentOptions (persistence model).
 */
export function compileDocument(json: JSONContent, extensions?: Extensions): DocumentOptions {
  return getDocxManager(extensions).compile(json);
}

/**
 * Fill in office-open's ECMA-376 schema defaults that a hand-built JSON lacks.
 *
 * A document constructed by hand (not via {@link parseDOCX}) carries no
 * `doc.attrs.styles` (docDefaults: body font/size/spacing + the built-in style
 * table) and no `doc.attrs.sectionProperties` (page size, margins, docGrid
 * linePitch). Those are injected only by office-open's `parseDocument`, which a
 * hand-built doc bypasses — so without them the editor has no body font, no
 * page geometry, and no document grid for snapToGrid to pitch against, and
 * rendering/pagination drift.
 *
 * Harvests the defaults by round-tripping an EMPTY document through office-open
 * (`generateDOCXSync` → `parseDOCX`) and shallow-merging the resulting
 * `doc.attrs` UNDER the input's. Only document-level attrs are touched —
 * content nodes (paragraphs/runs/marks) pass through verbatim, avoiding the
 * mark pollution a full-content round-trip would cause (a paragraph's default
 * run props leak onto its text as a textStyle mark). Keys already set on
 * `json.attrs` win, so a doc that already carries its own styles/section
 * properties (e.g. from `parseDOCX` or a prior `getJSON`) is left unchanged.
 */
export function normalizeDocument(json: JSONContent, extensions?: Extensions): JSONContent {
  const defaults = parseDOCX(generateDOCXSync({ type: "doc", content: [] }, { extensions }));
  const baseAttrs = (defaults.attrs ?? {}) as Record<string, unknown>;
  // A hand-built doc (e.g. parseHTML output) carries attrs keys with null
  // values (schema defaults) — those are "lacking", not overrides, so they
  // must not shadow the harvested defaults.
  const userAttrs = Object.fromEntries(
    Object.entries((json.attrs ?? {}) as Record<string, unknown>).filter(([, v]) => v != null),
  );
  return { ...json, attrs: { ...baseAttrs, ...userAttrs } };
}
