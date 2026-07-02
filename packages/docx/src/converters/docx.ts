import { encodeBase64 } from "@office-open/core";
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
  TableRowOptions,
  TableCellOptions,
  BorderOptions,
  LevelsOptions,
  NumberingOptions,
  OutputByType,
  OutputType,
  PackerOptions,
  StylesOptions,
  TableOfContentsOptions,
} from "@office-open/docx";
import { emojis, shortcodeToEmoji } from "@tiptap/extension-emoji";

import type { JSONContent } from "../core";
import * as blockquoteExt from "../extensions/blockquote";
import * as codeBlockExt from "../extensions/code-block";
import * as detailsExt from "../extensions/details";
import * as headingExt from "../extensions/heading";
import * as imageExt from "../extensions/image";
import * as mentionExt from "../extensions/mention";
import * as orderedListExt from "../extensions/ordered-list";
// Node renderDocx/parseDocx
import * as paragraphExt from "../extensions/paragraph";
// Mark renderDocx/parseDocx
import * as strikeExt from "../extensions/strike";
import * as tableExt from "../extensions/table";
import * as tableCellExt from "../extensions/table-cell";
import * as tableHeaderExt from "../extensions/table-header";
import * as tableRowExt from "../extensions/table-row";
import * as taskItemExt from "../extensions/task-item";
import * as textStyleExt from "../extensions/text-style";
import { alignmentFromCss } from "../extensions/utils";
import { prepareDocument, type PrepareStep } from "./prepare";
import { indexParagraphStyles, mergeTableStyleProps } from "./styles";

export type { DocumentOptions };

// ── Helpers ──

/** Remove keys with null/undefined values */
function cleanAttrs(attrs: Record<string, unknown>): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null && value !== undefined) result[key] = value;
  }
  return result;
}

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

/** True when a tblBorders object carries no REAL edge — every side is absent,
 *  none, or nil. office-open fills table.borders with all-`none` when the
 *  table's own <w:tblPr> defines no <w:tblBorders>, so this detects "the table
 *  has no borders of its own" to decide whether a referenced table style's
 *  borders should fill the gap. */
function allBordersNone(borders: unknown): boolean {
  if (!borders || typeof borders !== "object") return true;
  const b = borders as Record<string, BorderOptions | undefined>;
  return (["top", "bottom", "left", "right", "insideHorizontal", "insideVertical"] as const).every(
    (k) => {
      const v = b[k];
      return !v || v.style === "none" || v.style === "nil";
    },
  );
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
 * documentExtras pass-through so they aren't duplicated.
 */
const COMPILE_OWNED_KEYS = new Set<string>([
  "sections",
  "numbering",
  "styles",
  "background",
  ...CORE_PROPERTY_KEYS,
]);

/** Collect core properties present on DocumentOptions into a plain object. */
function extractCoreProperties(docOpts: DocumentOptions): DocxCoreProperties | null {
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

/** Merge consecutive text nodes with same marks */
function mergeTextNodes(nodes: JSONContent[]): JSONContent[] {
  const result: JSONContent[] = [];
  for (const node of nodes) {
    if (node.type === "text" && result.length > 0 && result[result.length - 1].type === "text") {
      const prev = result[result.length - 1];
      if (JSON.stringify(prev.marks) === JSON.stringify(node.marks)) {
        prev.text = (prev.text ?? "") + (node.text ?? "");
        continue;
      }
    }
    result.push({ ...node });
  }
  return result;
}

// ── Blockquote signature ──

// ── DocxManager ──

/**
 * Manages DOCX serialization (Tiptap JSON ↔ DocumentOptions).
 *
 * Each extension provides renderDocx/parseDocx for its own attrs mapping.
 * DocxManager handles tree walking, child assembly, and dispatching.
 */
export class DocxManager {
  // Numbering definitions accumulated during compile (ordered lists). Bullet
  // lists use office-open's built-in numbering and contribute nothing here.
  private numberingConfigs: { reference: string; levels: LevelsOptions[] }[] = [];
  // Per-orderedList instance counter — each list gets its own concrete instance
  // so independent lists number separately, even when they share an abstractNum
  // (same start).
  private orderedInstanceCounter = 0;
  // Styles table (styles.xml) carried through resolve() so a paragraph can
  // resolve a NUMERIC pStyle whose NAME is a heading (styleId "2" → name
  // "heading 1") into a heading node. office-open lifts only pStyle literals
  // that ARE HeadingLevels ("Heading1".."Title"); real DOCX files often use
  // numeric ids. Set per resolve(); compile never reads it.
  private resolveStyles: StylesOptions | undefined;
  // Reference → level-0 format/start, for classifying numbering paragraphs as
  // bullet vs ordered. Built once from docOpts and shared by every resolve path
  // that walks a block stream — section children, header/footer, and table cell
  // children (a cell is just another SectionChild[] stream). Set per resolve();
  // compile never reads it.
  private resolveNumberingLookup: Map<string, { format?: string; start?: number }> | undefined;

  compile(json: JSONContent): DocumentOptions {
    this.numberingConfigs = [];
    this.orderedInstanceCounter = 0;

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
        if (child) {
          if (Array.isArray(child)) currentChildren.push(...child);
          else currentChildren.push(child);
        }
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
          | { config?: { reference: string; levels: LevelsOptions[] }[] }
          | undefined
      )?.config ?? [];
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
      ...(numberingConfig.length > 0
        ? { numbering: { config: numberingConfig } as NumberingOptions }
        : {}),
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
        if (Array.isArray(child)) children.push(...child);
        else children.push(child);
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
    this.resolveNumberingLookup = this.buildNumberingLookup(docOpts);

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
        if (last?.type === "paragraph") {
          last.attrs = { ...last.attrs, ...sectAttrs };
        } else {
          // Section ends on a non-paragraph (table/etc.) — Word still needs the
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
      case "heading":
        return { paragraph: this.compileHeadingNode(node) };
      case "blockquote": {
        // blockquote → each child paragraph/heading gets a left indent + left
        // border (DOCX's native blockquote expression; no dedicated element).
        // Iterate ALL children — the previous `return` inside the loop dropped
        // every paragraph after the first.
        const items: SectionChild[] = [];
        for (const child of node.content ?? []) {
          if (child.type === "paragraph" || child.type === "heading") {
            const para =
              child.type === "heading"
                ? this.compileHeadingNode(child)
                : this.compileParagraphNode(child);
            const paraObj = (
              typeof para === "string" ? { text: para } : (para as Record<string, unknown>)
            ) as Record<string, unknown>;
            blockquoteExt.applyBlockquoteStyle(paraObj);
            items.push({ paragraph: paraObj as ParagraphOptions });
          }
        }
        return items.length > 0 ? items : null;
      }
      case "codeBlock": {
        // codeBlock → paragraph styled "Code" (extension owns style/font).
        // Children go through the shared inline path (handles `\n`→break +
        // inline marks), so no special-case child logic lives here.
        const opts = codeBlockExt.renderDocx(node) as Record<string, unknown>;
        const childList = this.compileInlineContent(node.content);
        if (childList.length > 0) opts.children = childList;
        return { paragraph: this.simplifyParagraph(opts) };
      }
      case "horizontalRule":
        return {
          paragraph: { thematicBreak: true } as Record<string, unknown> as ParagraphOptions,
        };
      case "table":
        return { table: this.compileTableNode(node) };
      case "image": {
        const imageRun = imageExt.renderDocx(node);
        if (!imageRun) return null;
        return { paragraph: { children: [imageRun] } };
      }
      case "bulletList":
      case "orderedList":
      case "taskList":
        return this.compileListFromNode(node, 0);
      case "details":
        return this.compileDetailsNode(node);
      case "tocField": {
        const options = (node.attrs?.options as TableOfContentsOptions | undefined) ?? {};
        const entries: SectionChild[] = [];
        for (const child of node.content ?? []) {
          const compiled = this.compileSectionChild(child);
          if (!compiled) continue;
          if (Array.isArray(compiled)) entries.push(...compiled);
          else entries.push(compiled);
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

  private compileParagraphNode(node: JSONContent): ParagraphOptions {
    const opts = paragraphExt.renderDocx(node) as Record<string, unknown>;
    const childList = this.compileInlineContent(node.content);
    if (childList.length > 0) opts.children = childList;
    return this.simplifyParagraph(opts);
  }

  private compileHeadingNode(node: JSONContent): ParagraphOptions {
    const opts = headingExt.renderDocx(node) as Record<string, unknown>;
    const childList = this.compileInlineContent(node.content);
    if (childList.length > 0) opts.children = childList;
    return this.simplifyParagraph(opts);
  }

  /** Simple text optimization: merge plain runs into text field */
  private simplifyParagraph(opts: Record<string, unknown>): ParagraphOptions {
    const children = opts.children as ParagraphChild[] | undefined;
    if (!children || children.length === 0) return opts as ParagraphOptions;

    const allSimpleText = children.every(
      (c) =>
        typeof c === "object" && c !== null && "text" in c && Object.keys(c as object).length === 1,
    );
    if (allSimpleText) {
      const combined = children.map((c) => (c as { text: string }).text).join("");
      delete opts.children;
      if (combined && Object.keys(opts).length === 0)
        return combined as unknown as ParagraphOptions;
      if (combined) (opts as Record<string, unknown>).text = combined;
    }

    return opts as ParagraphOptions;
  }

  private compileTableNode(node: JSONContent): TableOptions {
    const opts = tableExt.renderDocx(node) as Record<string, unknown>;
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
    const rows: Record<string, unknown>[] = [];

    // Track active vertical spans from previous rows. `borders` carries the
    // restart cell's tcBorders so a rebuilt continuation cell (vMerge continue)
    // re-emits them: Word draws the last continuation's bottom as the merged
    // region's bottom edge and each continuation's left/right as column
    // separators. Without it the region's borders collapse to the table-level
    // default (often "none") after a docx→json→docx round-trip — the merged
    // cells then render with missing borders.
    type ActiveSpan = {
      colStart: number;
      colspan: number;
      remainingRows: number;
      borders?: unknown;
    };
    let activeSpans: ActiveSpan[] = [];

    for (const rowNode of node.content ?? []) {
      if (rowNode.type !== "tableRow") continue;

      const rowOpts = tableRowExt.renderDocx(rowNode) as Record<string, unknown>;
      const pmCells = (rowNode.content ?? []).filter(
        (c) => c.type === "tableCell" || c.type === "tableHeader",
      );

      // Snapshot spans from previous rows for this row
      const currentSpans = [...activeSpans].sort((a, b) => a.colStart - b.colStart);
      const newSpans: ActiveSpan[] = [];
      const compiledCells: Record<string, unknown>[] = [];

      let colIdx = 0;
      let cellIdx = 0;
      let spanIdx = 0;

      // Interleave vMerge continuation cells with actual cells
      while (spanIdx < currentSpans.length || cellIdx < pmCells.length) {
        const nextSpanCol =
          spanIdx < currentSpans.length ? currentSpans[spanIdx].colStart : Infinity;

        if (colIdx >= nextSpanCol) {
          // Insert vMerge continuation cell at this column
          const span = currentSpans[spanIdx];
          compiledCells.push({
            verticalMerge: "continue",
            columnSpan: span.colspan,
            // Inherit the restart cell's tcBorders so the merged region keeps
            // its edges (see ActiveSpan above).
            ...(span.borders ? { borders: span.borders } : {}),
            children: [{ paragraph: "" }],
          });
          colIdx += span.colspan;
          spanIdx++;
        } else {
          // Compile and place the actual cell
          const cell = this.compileTableCellNode(
            pmCells[cellIdx],
            tableCellMargins,
            insideH,
            insideV,
          );
          const cs = (cell.columnSpan as number) ?? 1;
          const rs = (cell.rowSpan as number) ?? 1;

          if (rs > 1) {
            delete cell.rowSpan;
            cell.verticalMerge = "restart";
            newSpans.push({
              colStart: colIdx,
              colspan: cs,
              remainingRows: rs - 1,
              borders: cell.borders,
            });
          }

          compiledCells.push(cell);
          colIdx += cs;
          cellIdx++;
        }
      }

      rowOpts.cells = compiledCells;
      rows.push(rowOpts);

      // Update active spans for the next row
      for (const span of currentSpans) span.remainingRows--;
      activeSpans = [...currentSpans.filter((s) => s.remainingRows > 0), ...newSpans];
    }

    opts.rows = rows;

    // Compute tblGrid column widths and table width
    if (colCount > 0) {
      const { columnWidths, tableWidth } = this.computeColumnWidths(node, colCount, opts);
      opts.columnWidths = columnWidths;
      if (!opts.width) opts.width = tableWidth;
      if (!opts.layout) opts.layout = "autofit";
    }

    return opts as unknown as TableOptions;
  }

  /** Count grid columns from the first table row (summing colspan). */
  private getTableColumnCount(tableNode: JSONContent): number {
    for (const rowNode of tableNode.content ?? []) {
      if (rowNode.type !== "tableRow") continue;
      let count = 0;
      for (const cell of rowNode.content ?? []) {
        if (cell.type === "tableCell" || cell.type === "tableHeader") {
          count += (cell.attrs?.colspan as number) ?? 1;
        }
      }
      return count;
    }
    return 0;
  }

  /**
   * Compute tblGrid column widths (twips) from first-row cell colwidth attrs.
   * Returns columnWidths array and the appropriate tableWidth.
   */
  private computeColumnWidths(
    tableNode: JSONContent,
    colCount: number,
    _opts: Record<string, unknown>,
  ): { columnWidths: number[]; tableWidth: { size: number; type: string } } {
    const DEFAULT_CONTENT_TWIPS = 9026; // A4 with 1-inch margins (~15.9cm)

    // Prefer the table's columnWidths attr (tblGrid from DOCX round-trip): it
    // carries exact twips, avoiding the lossy twips→px→twips detour through
    // per-cell colwidth.
    const tableColWidths = tableNode.attrs?.columnWidths as number[] | null | undefined;
    if (tableColWidths && tableColWidths.length > 0) {
      const filled = tableColWidths.slice(0, colCount);
      while (filled.length < colCount) {
        filled.push(filled[filled.length - 1] ?? Math.floor(DEFAULT_CONTENT_TWIPS / colCount));
      }
      const total = filled.reduce((a, b) => a + b, 0);
      return { columnWidths: filled, tableWidth: { size: total, type: "dxa" } };
    }

    const widths: (number | null)[] = Array.from({ length: colCount }, () => null);

    // Collect explicit px widths from the first row's cells
    for (const rowNode of tableNode.content ?? []) {
      if (rowNode.type !== "tableRow") continue;
      let colIdx = 0;
      for (const cell of rowNode.content ?? []) {
        if (cell.type !== "tableCell" && cell.type !== "tableHeader") continue;
        const colspan = (cell.attrs?.colspan as number) ?? 1;
        const colwidth = cell.attrs?.colwidth as number[] | null;
        if (colwidth) {
          for (let i = 0; i < colspan && colIdx + i < colCount; i++) {
            const px = colwidth[i] ?? colwidth[0];
            if (px) widths[colIdx + i] = px * 15; // px → twips (96 DPI)
          }
        }
        colIdx += colspan;
      }
      break; // first row only
    }

    const hasExplicit = widths.some((w) => w !== null);

    if (!hasExplicit) {
      // No explicit widths → equal distribution, percentage table width
      const equal = Math.floor(DEFAULT_CONTENT_TWIPS / colCount);
      return {
        columnWidths: Array(colCount).fill(equal),
        tableWidth: { size: 5000, type: "pct" }, // 100%
      };
    }

    // Fill gaps with the average of explicit widths
    const explicit = widths.filter((w): w is number => w !== null);
    const avg = Math.floor(explicit.reduce((a, b) => a + b, 0) / explicit.length);
    const filled = widths.map((w) => w ?? avg);
    const total = filled.reduce((a, b) => a + b, 0);

    return {
      columnWidths: filled,
      tableWidth: { size: total, type: "dxa" },
    };
  }

  private compileTableCellNode(
    cellNode: JSONContent,
    tableMargins?: NonNullable<TableCellOptions["margins"]> | null,
    insideH?: BorderOptions | null,
    insideV?: BorderOptions | null,
  ): Record<string, unknown> {
    const cellOpts = (
      cellNode.type === "tableHeader"
        ? tableHeaderExt.renderDocx(cellNode)
        : tableCellExt.renderDocx(cellNode)
    ) as Record<string, unknown>;

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

    // Cell horizontal alignment (Tiptap base-extension `align` attr) is NOT an
    // OOXML cell property — <w:tcPr> has no horizontal alignment. Push it down to
    // each contained paragraph's `alignment` (the OOXML <w:jc>), unless a paragraph
    // already specifies its own alignment.
    // `attrs.align` is a CSS text-align value (left/center/right/justify). Map it
    // to an OOXML AlignmentType (justify → "both") for the paragraph <w:jc>.
    const cellAlign = alignmentFromCss((cellNode.attrs?.align as string | undefined) ?? null) as
      | ParagraphOptions["alignment"]
      | null;

    // A cell may contain ANY block (nested table/list/blockquote/codeBlock/…),
    // not just paragraphs. Route each child through the shared block compiler so
    // a nested list or table survives the round-trip — previously every child
    // was forced through compileParagraphNode, silently dropping non-paragraph
    // blocks into an empty paragraph.
    const cellChildren: SectionChild[] = [];
    const pushChild = (child: SectionChild) => {
      // Push the cell's `align` down to each paragraph's <w:jc> (see above).
      if (cellAlign && typeof child === "object" && child !== null && "paragraph" in child) {
        const p = child.paragraph;
        if (typeof p === "string") child.paragraph = { text: p, alignment: cellAlign };
        else if (p && typeof p === "object" && !(p as Record<string, unknown>).alignment)
          (p as Record<string, unknown>).alignment = cellAlign;
      }
      cellChildren.push(child);
    };
    for (const childNode of cellNode.content ?? []) {
      const compiled = this.compileSectionChild(childNode);
      if (compiled == null) continue;
      if (Array.isArray(compiled)) compiled.forEach(pushChild);
      else pushChild(compiled);
    }
    if (cellChildren.length > 0) cellOpts.children = cellChildren;

    return cellOpts;
  }

  private compileListFromNode(node: JSONContent, level: number): SectionChild[] | null {
    const items: SectionChild[] = [];
    const isOrdered = node.type === "orderedList";
    const isTask = node.type === "taskList";

    // Ordered lists need an abstractNum definition (decimal). The reference is
    // keyed by `start` so lists with the same start share one definition; each
    // list still gets a distinct instance (independent counting).
    const numbering = node.attrs?.numbering as string | undefined;
    let ordered: { reference: string; instance: number } | undefined;
    if (isOrdered && !numbering) ordered = this.registerOrderedNumbering(node);

    for (const listItem of node.content ?? []) {
      if (listItem.type !== "listItem" && listItem.type !== "taskItem") continue;
      const checked = Boolean(listItem.attrs?.checked);

      for (const child of listItem.content ?? []) {
        if (child.type === "paragraph" || child.type === "heading") {
          // Preserve heading level (heading → compileHeadingNode, not paragraph).
          const para =
            child.type === "heading"
              ? this.compileHeadingNode(child)
              : this.compileParagraphNode(child);
          const paraObj =
            typeof para === "string"
              ? ({ text: para } as Record<string, unknown>)
              : (para as Record<string, unknown>);

          if (numbering) {
            paraObj.numbering = { reference: numbering, level };
          } else if (ordered) {
            paraObj.numbering = { reference: ordered.reference, instance: ordered.instance, level };
          } else {
            paraObj.bullet = { level };
          }

          if (isTask) this.injectTaskCheckbox(paraObj, checked);

          items.push({ paragraph: paraObj as ParagraphOptions });
        } else if (
          child.type === "bulletList" ||
          child.type === "orderedList" ||
          child.type === "taskList"
        ) {
          // Nested list — recurse one level deeper. Nested list paragraphs are
          // emitted as siblings of this item's paragraph (DOCX flattens lists
          // to a single paragraph sequence; the `level` field carries depth).
          const nested = this.compileListFromNode(child, level + 1);
          if (nested) items.push(...nested);
        }
      }
    }
    return items.length > 0 ? items : null;
  }

  /**
   * Register (or reuse) an abstractNum for an ordered list's `start`, and
   * return a fresh instance so this list counts independently of other lists
   * that share the same definition.
   */
  private registerOrderedNumbering(node: JSONContent): { reference: string; instance: number } {
    const start = Number(node.attrs?.start ?? 1) || 1;
    let entry = this.numberingConfigs.find((c) => Number(c.levels[0]?.start ?? 1) === start);
    if (!entry) {
      entry = {
        reference: `${orderedListExt.ORDERED_REFERENCE_PREFIX}-${this.numberingConfigs.length + 1}`,
        levels: orderedListExt.buildOrderedLevels(start),
      };
      this.numberingConfigs.push(entry);
    }
    this.orderedInstanceCounter += 1;
    return { reference: entry.reference, instance: this.orderedInstanceCounter };
  }

  /**
   * Prepend an inline checkbox SDT to a task paragraph. The SDT is tagged
   * "docen-task" so resolve can tell task items apart from ordinary paragraphs
   * that happen to contain an SDT.
   */
  private injectTaskCheckbox(paraObj: Record<string, unknown>, checked: boolean): void {
    let existing: unknown[] = [];
    if (Array.isArray(paraObj.children)) {
      existing = paraObj.children as unknown[];
    } else if (typeof paraObj.text === "string") {
      if (paraObj.text) existing = [{ text: paraObj.text }];
      delete paraObj.text;
    }
    paraObj.children = [taskItemExt.createTaskCheckbox(checked), ...existing];
  }

  /**
   * details → block-level group-SDT. The summary paragraph is tagged with a
   * fixed style so resolve can split it back out; content blocks flatten in
   * after it. (No native collapse in DOCX — structure round-trips, the view
   * stays expanded.)
   */
  private compileDetailsNode(node: JSONContent): SectionChild {
    const sdtChildren: SectionChild[] = [];
    for (const child of node.content ?? []) {
      if (child.type === "detailsSummary") {
        const inline = this.compileInlineContent(child.content);
        const summaryPara: Record<string, unknown> = { style: detailsExt.DETAILS_SUMMARY_STYLE };
        if (inline.length > 0) summaryPara.children = inline;
        sdtChildren.push({ paragraph: summaryPara as ParagraphOptions });
      } else if (child.type === "detailsContent") {
        for (const block of child.content ?? []) {
          const compiled = this.compileSectionChild(block);
          if (!compiled) continue;
          if (Array.isArray(compiled)) sdtChildren.push(...compiled);
          else sdtChildren.push(compiled);
        }
      }
    }
    return {
      sdt: { properties: { tag: detailsExt.DETAILS_TAG, group: true }, children: sdtChildren },
    } as SectionChild;
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
          children.push({ break: 1 } as Record<string, unknown> as ParagraphChild);
          break;
        case "pageBreak":
          children.push({ pageBreak: true } as Record<string, unknown> as ParagraphChild);
          break;
        case "columnBreak":
          children.push({ columnBreak: true } as Record<string, unknown> as ParagraphChild);
          break;
        case "tab":
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
          const imageRun = imageExt.renderDocx(node);
          if (imageRun) children.push(imageRun);
          break;
        }
        case "wpgGroup": {
          const wpgGroup = node.attrs?.wpgGroup;
          if (wpgGroup) children.push({ wpgGroup } as unknown as ParagraphChild);
          break;
        }
        case "wpsShape": {
          // Editable text body: compile each content paragraph back to a
          // ParagraphOptions and reattach under wpsShape.children. Mirrors the
          // tocField compile (compileSectionChild → unwrap .paragraph).
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
        case "mention": {
          children.push(
            mentionExt.createMention(
              String(node.attrs?.id ?? ""),
              String(node.attrs?.label ?? ""),
            ) as ParagraphChild,
          );
          break;
        }
        case "emoji": {
          // DOCX has no emoji structure — emit the glyph as a plain text run,
          // resolving it from the same emoji dataset base renders from (compile
          // is headless, so the shortcode name must be looked up explicitly).
          // Falls back to the :name: shortcode if the dataset has no match;
          // resolve degrades DOCX text back to a plain text node.
          const name = String(node.attrs?.name ?? "");
          const glyph = name ? shortcodeToEmoji(name, emojis)?.emoji : undefined;
          const text = glyph ?? (name ? `:${name}:` : "");
          if (text) children.push({ text } as Record<string, unknown> as ParagraphChild);
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
      if (i > 0) children.push({ break: 1 } as ParagraphChild);
      if (segments[i]) this.compileTextRun(segments[i], node.marks, children);
    }
  }

  /** Emit a single run for `text` with all inline marks applied. */
  private compileTextRun(
    text: string,
    marks: JSONContent["marks"],
    children: ParagraphChild[],
  ): void {
    const runOpts: Record<string, unknown> = { text };

    for (const mark of marks ?? []) {
      switch (mark.type) {
        case "bold":
          runOpts.bold = true;
          break;
        case "italic":
          runOpts.italic = true;
          break;
        case "underline":
          runOpts.underline = { type: "single" };
          break;
        case "strike": {
          const strikeProps = strikeExt.renderDocx((mark.attrs ?? {}) as Record<string, unknown>);
          Object.assign(runOpts, strikeProps);
          break;
        }
        case "subscript":
          runOpts.subScript = true;
          break;
        case "superscript":
          runOpts.superScript = true;
          break;
        case "highlight":
          runOpts.highlight = mark.attrs?.color ?? "yellow";
          break;
        case "code":
          // rStyle "CodeChar" is the precise round-trip carrier; Consolas is a
          // visual fallback when styles.xml lacks the CodeChar definition.
          runOpts.style = "CodeChar";
          runOpts.font = "Consolas";
          break;
        case "textStyle": {
          const tsProps = textStyleExt.renderDocx((mark.attrs ?? {}) as Record<string, unknown>);
          Object.assign(runOpts, tsProps);
          break;
        }
        case "link": {
          const href = mark.attrs?.href as string | undefined;
          if (href) {
            const { text: _, ...runWithoutText } = runOpts;
            const linkChildren: (RunOptions | string)[] = [];
            if (text) linkChildren.push({ ...runWithoutText, text } as RunOptions);
            children.push({
              hyperlink: {
                link: href.startsWith("#") ? undefined : href,
                anchor: href.startsWith("#") ? href.slice(1) : undefined,
                children: linkChildren,
              },
            });
            return;
          }
          break;
        }
      }
    }

    children.push(runOpts as RunOptions);
  }

  // ── Resolve: DocumentOptions → Tiptap JSON ──

  private resolveSectionChild(child: SectionChild): JSONContent | null {
    if ("paragraph" in child) {
      return this.resolveParagraph(child.paragraph);
    }
    if ("table" in child) {
      return this.resolveTable(child.table as unknown as Record<string, unknown>);
    }
    if ("sdt" in child) {
      const sdt = (child as { sdt: { properties?: { tag?: string }; children?: SectionChild[] } })
        .sdt;
      if (sdt.properties?.tag === detailsExt.DETAILS_TAG) {
        return this.resolveDetailsSdt(sdt);
      }
      // Generic SDT (non-details content control) — carry verbatim.
      return this.resolvePassthrough(child);
    }
    if ("toc" in child) {
      return this.resolveToc(child.toc);
    }
    // rawXml (incl. aggregated TOC field), bookmarkStart/End, textbox, altChunk,
    // subDoc, customXml — no native Tiptap node. Carry the SectionChild verbatim
    // so the round-trip is byte-faithful.
    return this.resolvePassthrough(child);
  }

  /** Wrap an opaque SectionChild in a passthrough atom (attrs.data = JSON). */
  private resolvePassthrough(child: SectionChild): JSONContent {
    return { type: "passthrough", attrs: { data: JSON.stringify(child) } };
  }

  /**
   * Resolve a table of contents into an editable `tableOfContents` container:
   * `attrs.options` carries the field switches, `content` is the entry
   * paragraphs. Each entry's inner HYPERLINK field has content-less runs that
   * office-open parses as `null`; resolving the entries through
   * `resolveSectionChild` → `resolveParagraphChildren` drops those nulls, so the
   * TOC no longer reaches the generate path as an opaque blob of nulls (the
   * `stringifyRunInline(null).break` crash). When `entries` is absent/empty
   * (a fresh, unrendered TOC), keep the node valid for `content: "block+"` with
   * a placeholder empty paragraph.
   */
  private resolveToc(toc: TableOfContentsOptions & { alias?: string }): JSONContent {
    const { entries, ...options } = toc;
    const content: JSONContent[] = [];
    for (const entry of entries ?? []) {
      const node = this.resolveSectionChild(entry);
      if (!node) continue;
      if (Array.isArray(node)) content.push(...node);
      else content.push(node);
    }
    if (content.length === 0) content.push({ type: "paragraph" });
    const node: JSONContent = { type: "tocField", content };
    const cleanOptions = cleanAttrs(options as Record<string, unknown>);
    if (Object.keys(cleanOptions).length > 0) node.attrs = { options: cleanOptions };
    return node;
  }

  /**
   * Resolve a details group-SDT: the summary-style paragraph becomes
   * detailsSummary, the remaining blocks fold into detailsContent.
   */
  private resolveDetailsSdt(sdt: {
    properties?: { tag?: string };
    children?: SectionChild[];
  }): JSONContent {
    const content: JSONContent[] = [];
    let summary: JSONContent[] | null = null;
    for (const child of sdt.children ?? []) {
      if ("paragraph" in child) {
        const para = child.paragraph as ParagraphOptions;
        if (
          (para as unknown as Record<string, unknown>).style === detailsExt.DETAILS_SUMMARY_STYLE
        ) {
          summary = this.resolveInlineContent(para);
          continue;
        }
      }
      const node = this.resolveSectionChild(child);
      if (!node) continue;
      if (Array.isArray(node)) content.push(...node);
      else content.push(node);
    }
    const details: JSONContent = { type: "details", content: [] };
    if (summary !== null) details.content!.push({ type: "detailsSummary", content: summary });
    if (content.length > 0) details.content!.push({ type: "detailsContent", content });
    return details;
  }

  private resolveParagraph(opts: string | ParagraphOptions): JSONContent {
    const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;

    // horizontalRule: a paragraph reduced to a bottom border (thematicBreak)
    if (resolved.thematicBreak) {
      return { type: "horizontalRule" };
    }

    // codeBlock: paragraphs styled "Code"
    if (resolved.style === "Code") {
      return this.resolveCodeBlock(resolved);
    }

    // Detect heading: office-open's lifted `heading` literal, an explicit
    // outlineLevel, or a pStyle that names a heading style (directly, by
    // localized name, or via basedOn). See detectHeadingLevel for the order.
    const headingLevel = this.detectHeadingLevel(resolved);
    const nodeType = headingLevel ? "heading" : "paragraph";

    // Dispatch to extension parseDocx
    const attrs = headingLevel
      ? headingExt.parseDocx(resolved as unknown as Record<string, unknown>)
      : paragraphExt.parseDocx(resolved as unknown as Record<string, unknown>);

    // Heading7-9, numeric styleIds, and basedOn/outlineLevel-derived levels
    // aren't HeadingLevel literals, so parseDocx can't always derive `level`
    // from resolved.heading/style — stamp it from the detected headingLevel. The
    // real pStyle still rides on attrs.styleId (parseDocx carries resolved.style).
    if (headingLevel && attrs.level == null) attrs.level = headingLevel;

    // List paragraphs never reach here — resolveSectionChildren intercepts
    // them upstream and rebuilds the nested list tree.
    const content = this.resolveInlineContent(resolved);
    const cleanAttrsObj = cleanAttrs(attrs);

    const node: JSONContent = { type: nodeType };
    if (Object.keys(cleanAttrsObj).length > 0) node.attrs = cleanAttrsObj;
    if (content.length > 0) node.content = content;

    return node;
  }

  /** Heading level (1-9) for a paragraph, or undefined when it isn't a heading.
   *  DOCX marks a heading several ways, checked in priority order:
   *  1. office-open lifts a HeadingLevel pStyle ("Heading1".."Title") into `heading`.
   *  2. An explicit `outlineLevel` (0-8 → 1-9) — Word's outline/TOC key off this
   *     even without a heading pStyle; the Heading1-9 styles carry outlineLvl 0-8.
   *  3. A pStyle that names a heading style: directly ("Heading7", which stays on
   *     `style` because office-open's HeadingLevel type caps at 6), by localized
   *     NAME ("heading 1"/"标题 1"), or via the `basedOn` chain (a custom style
   *     "MyTitle" basedOn="Heading1"). `heading` and `style` carry the same pStyle.
   *  `outlineLevel` is read loosely — office-open's public type omits the field
   *  even though it round-trips (w:outlineLvl) at runtime. */
  private detectHeadingLevel(resolved: ParagraphOptions): number | undefined {
    if (resolved.heading) {
      const lvl = HEADING_LEVEL_MAP[resolved.heading];
      if (lvl) return lvl;
    }
    const outline = (resolved as { outlineLevel?: number }).outlineLevel;
    if (typeof outline === "number" && outline >= 0 && outline <= 8) {
      return (outline + 1) as 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
    }
    const styleId = resolved.style;
    if (!styleId || !this.resolveStyles) return undefined;
    const byId = indexParagraphStyles(this.resolveStyles);
    const visited = new Set<string>();
    let curId: string | undefined = styleId;
    while (curId && !visited.has(curId)) {
      visited.add(curId);
      if (HEADING_LEVEL_MAP[curId]) return HEADING_LEVEL_MAP[curId];
      const style = byId.get(curId);
      if (!style) break;
      const lvl = headingLevelFromName(style.name);
      if (lvl) return lvl;
      curId = style.basedOn ?? undefined;
    }
    return undefined;
  }

  /** reference → level-0 format/start, for classifying numbering paragraphs. */
  private buildNumberingLookup(
    docOpts: DocumentOptions,
  ): Map<string, { format?: string; start?: number }> {
    const lookup = new Map<string, { format?: string; start?: number }>();
    const config = (
      docOpts as { numbering?: { config?: { reference: string; levels: LevelsOptions[] }[] } }
    ).numbering?.config;
    if (config) {
      for (const entry of config) {
        const lvl0 = entry.levels[0];
        lookup.set(entry.reference, { format: lvl0?.format, start: lvl0?.start });
      }
    }
    return lookup;
  }

  /**
   * Walk a SectionChild[] block stream — a section's body, a header/footer
   * slot, or a table cell's children (a cell is just another block stream) —
   * grouping consecutive list paragraphs into nested Tiptap lists. Non-list
   * children resolve individually. DOCX flattens lists to a paragraph sequence
   * (depth carried by `level`); this rebuilds the tree.
   */
  private resolveSectionChildren(children: SectionChild[]): JSONContent[] {
    const content: JSONContent[] = [];
    let i = 0;
    while (i < children.length) {
      // Bind to a local: TS won't `in`-narrow an indexed read (`children[i]`)
      // across two accesses, so the narrowing needs a stable binding. The
      // `typeof` guard also rejects the plain-string paragraph shorthand,
      // which is never a list item.
      const child = children[i];
      const firstPara = "paragraph" in child ? child.paragraph : null;
      const firstInfo =
        firstPara && typeof firstPara !== "string" ? this.detectList(firstPara) : null;

      if (!firstInfo) {
        // Blockquote: consecutive paragraphs carrying the blockquote signature
        // (left indent + left border). DOCX has no blockquote element, so
        // compile stamps this signature; rebuild the container here.
        if (firstPara && typeof firstPara !== "string" && this.detectBlockquote(firstPara)) {
          const group: ParagraphOptions[] = [];
          while (i < children.length) {
            const member = children[i];
            if (!("paragraph" in member)) break;
            const para = member.paragraph;
            if (typeof para === "string" || !this.detectBlockquote(para)) break;
            group.push(para);
            i++;
          }
          content.push(this.buildBlockquote(group));
          continue;
        }
        const node = this.resolveSectionChild(child);
        if (node) content.push(node);
        i++;
        continue;
      }

      // Collect the run of consecutive list paragraphs. A non-paragraph or a
      // plain-text paragraph ends the run — plain text is never a list item.
      const group: { para: ParagraphOptions; info: ListInfo }[] = [];
      while (i < children.length) {
        const member = children[i];
        if (!("paragraph" in member)) break;
        const para = member.paragraph;
        if (typeof para === "string") break;
        const info = this.detectList(para);
        if (!info) break;
        group.push({ para, info });
        i++;
      }
      content.push(...this.buildListTree(group));
    }
    return content;
  }

  /** Classify a paragraph as a list item, or null if it isn't one. */
  private detectList(para: ParagraphOptions): ListInfo | null {
    const p = para as unknown as Record<string, unknown>;
    const numbering = p.numbering as { reference?: string; level?: number } | undefined;
    const bullet = p.bullet as { level?: number } | undefined;

    let kind: "bullet" | "ordered";
    let level: number;
    let reference: string | undefined;
    let start: number | undefined;

    if (numbering) {
      reference = numbering.reference;
      level = numbering.level ?? 0;
      const cfg = reference ? this.resolveNumberingLookup?.get(reference) : undefined;
      // A config whose format isn't "bullet" → ordered; otherwise this is the
      // built-in default-bullet numbering (parse may tag numId=1 as numbering
      // when its abstractNum resolves), so degrade to bullet.
      if (cfg && cfg.format && cfg.format !== "bullet") {
        kind = "ordered";
        start = cfg.start;
      } else {
        kind = "bullet";
        // Keep the source reference: a custom bullet abstractNum (e.g. a
        // Wingdings glyph) needs its original definition to round-trip the
        // marker; buildListTree carries it on the list node for compile.
      }
    } else if (bullet) {
      kind = "bullet";
      level = bullet.level ?? 0;
    } else {
      return null;
    }

    // Task items carry a leading inline checkbox SDT tagged "docen-task".
    const first = (p.children as unknown[] | undefined)?.[0];
    const isTask = taskItemExt.isTaskCheckbox(first);

    return {
      kind: isTask ? "task" : kind,
      level,
      reference,
      start,
      checked: taskItemExt.readCheckboxState(first),
    };
  }

  /**
   * Rebuild nested Tiptap lists from a flat run of list paragraphs. Stack-based:
   * each frame is an active list at a given depth; the `key` (level:type:
   * reference) decides whether a paragraph continues the top list, starts a
   * nested list, or splits off a new sibling list.
   */
  private buildListTree(group: { para: ParagraphOptions; info: ListInfo }[]): JSONContent[] {
    const topLevel: JSONContent[] = [];
    const stack: {
      level: number;
      key: string;
      listNode: JSONContent;
      currentItem: JSONContent;
    }[] = [];

    for (const { para, info } of group) {
      const listType =
        info.kind === "ordered" ? "orderedList" : info.kind === "task" ? "taskList" : "bulletList";
      const itemType = info.kind === "task" ? "taskItem" : "listItem";
      const key = `${info.level}:${listType}:${info.reference ?? ""}`;

      // Pop frames that are deeper than this item, or at the same depth but a
      // different list (level/type/reference change → new list).
      while (stack.length > 0) {
        const top = stack[stack.length - 1];
        if (top.level > info.level || (top.level === info.level && top.key !== key)) {
          stack.pop();
          continue;
        }
        break;
      }

      const itemPara = this.resolveListItemParagraph(para, info);
      const newItem: JSONContent = { type: itemType, content: [itemPara] };
      if (itemType === "taskItem") newItem.attrs = { checked: info.checked };

      const top = stack[stack.length - 1];
      if (top && top.level === info.level && top.key === key) {
        // Same list continues — append a new item.
        (top.listNode.content as JSONContent[]).push(newItem);
        top.currentItem = newItem;
      } else {
        // New list (top-level or nested under the current item).
        const newList: JSONContent = { type: listType, content: [newItem] };
        const listAttrs: Record<string, unknown> = {};
        // Only level-0 ordered lists carry `start`; deeper levels restart at 1.
        if (
          listType === "orderedList" &&
          info.level === 0 &&
          typeof info.start === "number" &&
          info.start !== 1
        ) {
          listAttrs.start = info.start;
        }
        // Carry the source abstractNum reference so the marker round-trips.
        if (info.reference) listAttrs.numbering = info.reference;
        if (Object.keys(listAttrs).length > 0) newList.attrs = listAttrs;
        if (top) {
          (top.currentItem.content as JSONContent[]).push(newList);
        } else {
          topLevel.push(newList);
        }
        stack.push({ level: info.level, key, listNode: newList, currentItem: newItem });
      }
    }

    return topLevel;
  }

  /** Classify a paragraph as a blockquote member by its signature. */
  private detectBlockquote(para: ParagraphOptions): boolean {
    const p = para as unknown as Record<string, unknown>;
    const indent = p.indent as { left?: number } | undefined;
    const border = p.border as { left?: Record<string, unknown> } | undefined;
    if (!indent || indent.left !== blockquoteExt.BLOCKQUOTE_INDENT_LEFT) return false;
    const bl = border?.left;
    if (!bl) return false;
    const sig = blockquoteExt.BLOCKQUOTE_BORDER;
    return (
      bl.style === sig.style &&
      bl.size === sig.size &&
      bl.space === sig.space &&
      bl.color === sig.color
    );
  }

  /**
   * Rebuild a blockquote node from a run of signature-carrying paragraphs,
   * stripping the indent/border signature so child paragraphs render clean.
   */
  private buildBlockquote(group: ParagraphOptions[]): JSONContent {
    const content: JSONContent[] = [];
    for (const para of group) {
      const node = this.resolveParagraph(para);
      const attrs = node.attrs as Record<string, unknown> | undefined;
      if (attrs) {
        if (attrs.indent) {
          const indent = { ...(attrs.indent as object) } as Record<string, unknown>;
          delete indent.left;
          attrs.indent = Object.keys(indent).length > 0 ? indent : undefined;
        }
        if (attrs.border) {
          const border = { ...(attrs.border as object) } as Record<string, unknown>;
          delete border.left;
          attrs.border = Object.keys(border).length > 0 ? border : undefined;
        }
        const cleaned = cleanAttrs(attrs);
        if (Object.keys(cleaned).length > 0) node.attrs = cleaned;
        else delete node.attrs;
      }
      content.push(node);
    }
    return { type: "blockquote", content };
  }

  /**
   * Resolve a list-item paragraph to a Tiptap paragraph/heading node, stripping
   * the list marker (bullet/numbering) and the leading task checkbox — those
   * are expressed at the list/item level, not inside the paragraph.
   */
  private resolveListItemParagraph(para: ParagraphOptions, info: ListInfo): JSONContent {
    const resolved = typeof para === "string" ? ({ text: para } as ParagraphOptions) : para;
    const headingLevel = this.detectHeadingLevel(resolved);
    const nodeType = headingLevel ? "heading" : "paragraph";

    const attrs = headingLevel
      ? headingExt.parseDocx(resolved as unknown as Record<string, unknown>)
      : paragraphExt.parseDocx(resolved as unknown as Record<string, unknown>);
    // See resolveParagraph: derived levels aren't always literals, so stamp level.
    if (headingLevel && attrs.level == null) attrs.level = headingLevel;

    // Task: drop the leading checkbox SDT (its state lives in taskItem.attrs).
    const stripped = info.kind === "task" ? this.stripTaskCheckbox(resolved) : resolved;
    const content = this.resolveInlineContent(stripped);

    const node: JSONContent = { type: nodeType };
    const cleanAttrsObj = cleanAttrs(attrs);
    if (Object.keys(cleanAttrsObj).length > 0) node.attrs = cleanAttrsObj;
    if (content.length > 0) node.content = content;
    return node;
  }

  /** Return a copy of `para` with its leading docen-task checkbox SDT removed. */
  private stripTaskCheckbox(para: ParagraphOptions): ParagraphOptions {
    const children = (para as unknown as Record<string, unknown>).children;
    if (Array.isArray(children) && children.length > 0 && taskItemExt.isTaskCheckbox(children[0])) {
      return { ...(para as object), children: children.slice(1) } as ParagraphOptions;
    }
    return para;
  }

  private resolveCodeBlock(opts: ParagraphOptions): JSONContent {
    // Reassemble code: break → "\n" (merged into text), runs keep their marks.
    const children = opts.children as (ParagraphChild | string)[] | undefined;
    const content: JSONContent[] = [];
    if (children) {
      for (const child of children) {
        if (typeof child === "string") {
          if (child) content.push({ type: "text", text: child });
        } else if (typeof child === "object" && child !== null) {
          if ("break" in child) {
            const prev = content[content.length - 1];
            if (prev && prev.type === "text") prev.text = (prev.text ?? "") + "\n";
            else content.push({ type: "text", text: "\n" });
          } else if ("text" in child) {
            const marks = this.resolveMarks(child as RunOptions);
            const textNode: JSONContent = {
              type: "text",
              text: (child as { text: string }).text,
            };
            if (marks) textNode.marks = marks;
            content.push(textNode);
          }
        }
      }
    } else if (opts.text) {
      content.push({ type: "text", text: opts.text });
    }
    const node: JSONContent = { type: "codeBlock" };
    if (content.length > 0) node.content = content;
    return node;
  }

  private resolveTable(tableOpts: Record<string, unknown>): JSONContent {
    const attrs = tableExt.parseDocx(tableOpts);
    const rows = (tableOpts.rows ?? []) as Record<string, unknown>[];
    const content: JSONContent[] = [];

    // Pull the referenced table style's tblBorders/tblCellMar in: office-open
    // leaves table.borders/cellMargin reflecting only the table's own tblPr, so
    // a "Table Grid" table (borders defined in the style) would render no grid
    // without this. The table's own real borders win; the style fills the gap
    // when the table's are all none/nil.
    const styleProps = mergeTableStyleProps(
      this.resolveStyles?.tableStyles,
      (tableOpts.style as string | undefined) ?? null,
    );
    if (styleProps.borders && allBordersNone(tableOpts.borders)) {
      tableOpts = { ...tableOpts, borders: styleProps.borders };
    }
    if (styleProps.cellMargin && tableOpts.cellMargin == null && tableOpts.margins == null) {
      tableOpts = { ...tableOpts, cellMargin: styleProps.cellMargin };
    }

    // Table-level default cell insets (w:tblCellMar). office-open exposes them
    // on both `cellMargin` and `margins`; a cell inherits them unless it carries
    // its own tcMar. Push the default onto cells without tcMar so render
    // (renderTableCellStyles) and the paginator (cellVerticalOverhead) read ONE
    // effective source (cell.attrs.margins) instead of each falling back to the
    // table. compileTableCellNode drops a cell tcMar equal to this default to
    // keep the regenerated docx in its table-level form (near-identity round-trip).
    const tableCellMargins = (tableOpts.cellMargin ?? tableOpts.margins ?? null) as NonNullable<
      TableCellOptions["margins"]
    > | null;
    // Table-level inside grid lines (tblBorders.insideHorizontal/insideVertical).
    // In CSS border-collapse the interior grid belongs to cells, not the <table>
    // element, so a REAL insideH/V is pushed onto cell sides lacking their own
    // tcBorder (below). none/nil is skipped — a table that merely LACKS inner
    // grid lines must not have `border:none` stamped on every cell, or it would
    // suppress the editor's Table-Grid fallback default for borderless tables.
    // Edge cells' outer sides overlap the table's own border under
    // border-collapse (thicker wins), matching OOXML outer-vs-inner semantics.
    const tableBorders = (tableOpts.borders ?? null) as {
      insideHorizontal?: BorderOptions;
      insideVertical?: BorderOptions;
    } | null;
    const realBorder = (bd: unknown): BorderOptions | null => {
      const b = bd as BorderOptions | null | undefined;
      return b && b.style && b.style !== "none" && b.style !== "nil" ? b : null;
    };
    const insideH = realBorder(tableBorders?.insideHorizontal);
    const insideV = realBorder(tableBorders?.insideVertical);

    // DOCX encodes a row span as a `restart` cell followed by N empty `continue`
    // cells; ProseMirror uses one cell with rowspan = N+1. Track open spans per
    // column so each continuation cell increments the owning cell's `rowspan`,
    // which compileTableNode reads back to rebuild the continuation cells.
    let activeSpans = new Map<number, JSONContent>();

    for (const row of rows) {
      const rowAttrs = tableRowExt.parseDocx(row as Partial<TableRowOptions>);
      // gridAfter/widthAfter are rebuilt as trailing placeholder cells below
      // (PM requires every row's cell-span sum to equal the column count, so a
      // gridAfter row needs explicit empty cells or fixTables inserts the filler
      // at the START). Drop them from rowAttrs so compile doesn't re-emit
      // row.gridAfter on top of those cells — that double-counts and the row
      // widens by N columns every docx→json→docx round-trip.
      delete rowAttrs.gridAfter;
      delete rowAttrs.widthAfter;
      const cells = (row.cells ?? []) as Record<string, unknown>[];
      const cellNodes: JSONContent[] = [];
      const nextActiveSpans = new Map<number, JSONContent>();
      let colIdx = 0;

      for (const cell of cells) {
        const cellColspan = (cell.columnSpan as number) ?? 1;
        const vMerge = cell.verticalMerge as string | undefined;

        if (vMerge === "continue") {
          // Column owned by a cell above — bump its rowspan and carry the span
          // forward so further continuation cells keep counting.
          const owner = activeSpans.get(colIdx);
          if (owner) {
            const ownerAttrs = (owner.attrs ??= {});
            ownerAttrs.rowspan = ((ownerAttrs.rowspan as number) ?? 1) + 1;
            for (let c = colIdx; c < colIdx + cellColspan; c++) nextActiveSpans.set(c, owner);
          }
          colIdx += cellColspan;
          continue;
        }

        const isHeader = row.tableHeader as boolean;
        const cellAttrs = isHeader
          ? tableHeaderExt.parseDocx(cell as Partial<TableCellOptions>)
          : tableCellExt.parseDocx(cell as Partial<TableCellOptions>);

        // `rowspan` (recovered below) drives compile-time vMerge — drop the
        // OOXML marker so it doesn't round-trip back verbatim.
        delete cellAttrs.verticalMerge;

        // Effective cell margins: a cell's own tcMar wins, else inherit the
        // table's tblCellMar default (resolved once here for render + measure).
        if (!cellAttrs.margins && tableCellMargins) cellAttrs.margins = tableCellMargins;

        // Effective cell borders: a cell's own tcBorder per side wins, else
        // inherit the table's inside grid lines (insideH on top/bottom, insideV
        // on left/right) so the interior grid renders under border-collapse.
        // compileTableCellNode drops a side equal to the table's insideH/V to
        // keep the round-trip near-identity.
        if (insideH || insideV) {
          if (!cellAttrs.borders) cellAttrs.borders = {};
          const b = cellAttrs.borders as Record<string, BorderOptions | undefined>;
          if (insideH && !b.top) b.top = insideH;
          if (insideH && !b.bottom) b.bottom = insideH;
          if (insideV && !b.left) b.left = insideV;
          if (insideV && !b.right) b.right = insideV;
        }

        // A cell is just another SectionChild[] block stream — same as a
        // section body or a header/footer slot — so resolve it through the same
        // path. That regroups consecutive numbering/bullet paragraphs into list
        // nodes and keeps nested tables/lists structurally intact on import.
        const cellChildren = (cell.children ?? []) as SectionChild[];
        const cellContent: JSONContent[] = this.resolveSectionChildren(cellChildren);

        const cellType = isHeader ? "tableHeader" : "tableCell";
        const cellNode: JSONContent = { type: cellType };
        if (Object.keys(cellAttrs).length > 0) cellNode.attrs = cleanAttrs(cellAttrs);
        // An empty cell still needs content to satisfy the tableCell/tableHeader
        // `block+` schema. A content-less cell reaches the doc via fromJSON (which
        // skips validation), but prosemirror-tables' fixTables runs setNodeMarkup
        // on every table during appendTransaction, and setNodeMarkup re-validates
        // the node content — throwing "Invalid content for node type tableCell".
        // That throw aborts the paginator's reflow transaction, so the document
        // never re-pages (every block piles on page 0). Backfill an empty
        // paragraph; OOXML likewise requires a <w:p> in every <w:tc>.
        if (cellContent.length > 0) cellNode.content = cellContent;
        else cellNode.content = [{ type: "paragraph" }];

        if (vMerge === "restart") {
          // rowspan is finalized when continuation cells arrive below; register
          // the node so they can find and increment it.
          for (let c = colIdx; c < colIdx + cellColspan; c++) nextActiveSpans.set(c, cellNode);
        }

        cellNodes.push(cellNode);
        colIdx += cellColspan;
      }

      // OOXML gridAfter (w:gridAfter + widthAfter): N trailing grid columns this
      // row leaves uncovered. ProseMirror requires every row's cell-span sum to
      // equal the column count; without explicit trailing cells fixTables fills
      // the gap — and for a row whose only real cell is a leading gridSpan (e.g.
      // a header row: 1 cell spanning 2 + gridAfter 1) it inserts the filler
      // at the START, shoving the real cell right onto narrower columns (wrong
      // width + off-center). Emit explicit empty trailing cells so real cells
      // keep their left positions.
      const gridAfter = (row.gridAfter as number) ?? 0;
      if (gridAfter > 0) {
        const trailingType = (row.tableHeader as boolean) ? "tableHeader" : "tableCell";
        // gridAfter cells are empty trailing grid columns (no content). Give them
        // nil borders on every side so renderTableCellStyles emits border:none —
        // otherwise they pick up the Table-Grid default and draw a stray vertical
        // line at the row's right edge, showing up as an empty cell.
        const nilBorders = {
          top: { style: "nil" },
          right: { style: "nil" },
          bottom: { style: "nil" },
          left: { style: "nil" },
        };
        for (let c = 0; c < gridAfter; c++)
          cellNodes.push({
            type: trailingType,
            attrs: { borders: nilBorders },
            content: [{ type: "paragraph" }],
          });
        colIdx += gridAfter;
      }

      activeSpans = nextActiveSpans;

      const rowNode: JSONContent = { type: "tableRow" };
      if (Object.keys(rowAttrs).length > 0) rowNode.attrs = cleanAttrs(rowAttrs);
      if (cellNodes.length > 0) rowNode.content = cellNodes;

      content.push(rowNode);
    }

    const node: JSONContent = { type: "table" };
    if (Object.keys(attrs).length > 0) node.attrs = cleanAttrs(attrs);
    if (content.length > 0) node.content = content;

    return node;
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
        if (resolved) nodes.push(...(Array.isArray(resolved) ? resolved : [resolved]));
      }
    }
    return nodes;
  }

  private resolveParagraphChild(child: ParagraphChild): JSONContent | JSONContent[] | null {
    // `<w:r><w:tab/></w:r>` → office-open ParagraphChild `{ tab: true }` (a pure
    // tab run, e.g. between a TOC entry's title and page number). Turn it into a
    // `tab` inline atom so the leader can render; otherwise it fell through to
    // inlinePassthrough (hidden, no leader) and mergeTextNodes collapsed the
    // adjacent title/page-number text together.
    if ("tab" in child) {
      return { type: "tab" };
    }
    if ("text" in child || "children" in child || "break" in child) {
      return this.resolveRun(child as RunOptions);
    }
    if ("image" in child) {
      return this.resolveImage(child.image as unknown as Record<string, unknown>);
    }
    if ("wpgGroup" in child) {
      // Drawing group (wpg): opaque round-trip — full WpgGroupRunOptions rides on
      // the wpgGroup node; the editor doesn't model the group interior.
      return { type: "wpgGroup", attrs: { wpgGroup: child.wpgGroup } };
    }
    if ("wpsShape" in child) {
      // Standalone floating text box (wp:anchor > wps:wsp, NOT inside a wpg
      // group). The shape's text body (children: (ParagraphOptions | string)[])
      // becomes PM content (one node per paragraph); geometry/styling ride on
      // attrs.wpsShape. Mirrors resolveToc: split body → content, keep the rest.
      const ws = child.wpsShape;
      const content: JSONContent[] = [];
      if (ws?.children) {
        for (const para of ws.children) {
          if (typeof para !== "object" || para === null) {
            const node = this.resolveParagraph(para);
            if (node) content.push(node);
            continue;
          }
          // DrawingML defRPr (para.run) is the default run-properties for the
          // box's runs, NOT the OOXML ¶-mark rPr. Merge it into each run
          // (matching the prior atom renderWpsText: {...para.run, ...r}), then
          // drop it from the paragraph (run: undefined): paragraph.ts renders
          // attrs.run.size as ¶-mark line-height (renderParagraphStyles
          // markLineHeight), which would override the box's grid line-height —
          // but defRPr is a run default, not a ¶ mark. PDF measures 27.5pt
          // (the body grid); dropping defRPr lets the paragraph inherit it.
          // Round-trip safe — runs carry the full rPr, so compile emits
          // per-run rPr and Word renders identically.
          const defRPr = (para.run as Record<string, unknown> | undefined) ?? {};
          const children = Array.isArray(para.children)
            ? para.children.map((c) =>
                typeof c !== "object" || c === null
                  ? { ...defRPr, text: c as string }
                  : { ...defRPr, ...(c as object) },
              )
            : undefined;
          const node = this.resolveParagraph({
            ...para,
            run: undefined,
            ...(children ? { children } : {}),
          });
          if (node) content.push(node);
        }
      }
      if (content.length === 0) content.push({ type: "paragraph" });
      const { children: _omit, ...geometry } = ws ?? {};
      const node: JSONContent = { type: "wpsShape", content };
      const cleanGeometry = cleanAttrs(geometry as Record<string, unknown>);
      if (Object.keys(cleanGeometry).length > 0) node.attrs = { wpsShape: cleanGeometry };
      return node;
    }
    if ("sdt" in child) {
      return this.resolveInlineSdt(child);
    }
    if ("hyperlink" in child) {
      return this.resolveHyperlink(
        child.hyperlink as {
          link?: string;
          anchor?: string;
          tooltip?: string;
          children?: (RunOptions | string)[];
        },
      );
    }
    if ("pageBreak" in child) {
      return { type: "pageBreak" };
    }
    if ("columnBreak" in child) {
      return { type: "columnBreak" };
    }
    // Unrecognized inline child (bookmark/range markers, proofErr, track-change
    // markers, …) — carry verbatim via inlinePassthrough so the round-trip stays
    // byte-faithful (mirrors block-level resolvePassthrough).
    return { type: "inlinePassthrough", attrs: { data: JSON.stringify(child) } };
  }

  /** Resolve an inline SDT (mention carrier; other inline SDTs unsupported). */
  private resolveInlineSdt(child: ParagraphChild): JSONContent | null {
    if (mentionExt.isMention(child)) {
      const { id, label } = mentionExt.readMention(child);
      return { type: "mention", attrs: { id, label } };
    }
    return null;
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
          if ("pageBreak" in c) {
            flushText();
            nodes.push({ type: "pageBreak" });
          } else if ("columnBreak" in c) {
            flushText();
            nodes.push({ type: "columnBreak" });
          } else if ("break" in c) {
            flushText();
            nodes.push({ type: "hardBreak" });
          } else if ("tab" in c) {
            flushText();
            nodes.push({ type: "tab" });
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

    if (opts.bold) marks.push({ type: "bold" });
    if (opts.italic) marks.push({ type: "italic" });
    if (opts.underline) marks.push({ type: "underline" });
    if (opts.strike) {
      const strikeAttrs = strikeExt.parseDocx(opts);
      marks.push({ type: "strike", attrs: { doubleStrike: strikeAttrs.doubleStrike ?? null } });
    }
    if (opts.subScript) marks.push({ type: "subscript" });
    if (opts.superScript) marks.push({ type: "superscript" });
    if (opts.highlight) marks.push({ type: "highlight", attrs: { color: opts.highlight } });

    if (opts.style === "CodeChar") {
      marks.push({ type: "code" });
    }

    const textStyleAttrs = textStyleExt.parseDocx(opts);
    // parseDocx passes RunStylePropertiesOptions through verbatim (size/color/font/
    // characterSpacing/rightToLeft/…). The `code` mark is carried by rStyle
    // "CodeChar" (resolved above) and renders Consolas via its own font — drop
    // that font from text-style attrs to avoid duplication. color normalization
    // happens in renderHTML via normalizeColorToHex.
    if (opts.style === "CodeChar") {
      delete textStyleAttrs.font;
      // CodeChar is carried by the `code` mark above; don't also stamp it as a
      // textStyle styleId (would double-apply the character style on compile).
      delete textStyleAttrs.styleId;
    }

    if (Object.keys(textStyleAttrs).length > 0) {
      marks.push({ type: "textStyle", attrs: textStyleAttrs });
    }

    return marks.length > 0 ? marks : undefined;
  }

  private resolveImage(imageOpts: Record<string, unknown>): JSONContent {
    const attrs = imageExt.parseDocx(imageOpts);

    // Image data → data URL (encodeBase64 handles platform dispatch + stack guard).
    const data = imageOpts.data as Uint8Array | undefined;
    const type = imageOpts.type as string | undefined;
    if (data && type) {
      const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
      attrs.src = `data:image/${type};base64,${encodeBase64(bytes)}`;
    }

    return { type: "image", attrs };
  }

  private resolveHyperlink(hyperlink: {
    link?: string;
    anchor?: string;
    tooltip?: string;
    children?: (RunOptions | string)[];
  }): JSONContent | null {
    const href = hyperlink.link ?? (hyperlink.anchor ? `#${hyperlink.anchor}` : "");
    if (!href) return null;

    const content = this.resolveParagraphChildren(
      (hyperlink.children ?? []).map((c) => c as ParagraphChild),
    );

    if (content.length > 0) {
      const merged = mergeTextNodes(content);
      for (const node of merged) {
        if (node.type === "text") {
          node.marks = [
            ...(node.marks ?? []),
            {
              type: "link",
              attrs: {
                href,
                // Internal anchor (#bookmark, e.g. TOC entry jumps) must stay in
                // the current window so the in-page jump resolves; only external
                // links open a new tab.
                target: href.startsWith("#") ? null : "_blank",
                rel: "noopener noreferrer nofollow",
                class: null,
                title: hyperlink.tooltip ?? null,
              },
            },
          ];
        }
      }
      return merged;
    }

    return null;
  }
}

// ── List reconstruction ──

interface ListInfo {
  kind: "bullet" | "ordered" | "task";
  level: number;
  reference?: string;
  start?: number;
  checked: boolean;
}

// ── Heading level map ──

const HEADING_LEVEL_MAP: Record<string, 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Heading7: 7,
  Heading8: 8,
  Heading9: 9,
  Title: 1,
};

/** Heading level (1-9) from a localized style NAME: "heading 1"/"标题 1" → 1,
 *  "title" → 1. office-open's built-in names are English ("heading 1"), but
 *  zh-CN Word labels the same styles "标题 1"; both map to the same level. */
function headingLevelFromName(name: string | undefined): number | undefined {
  if (!name) return undefined;
  const m = /^heading\s+(\d)$/i.exec(name) ?? /^标题\s*(\d)$/.exec(name);
  if (m) {
    const lvl = Number(m[1]);
    if (lvl >= 1 && lvl <= 9) return lvl as 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  }
  return /^title$/i.test(name) ? 1 : undefined;
}

// ── Standalone functions (backward compat) ──

const defaultManager = new DocxManager();

/**
 * Parse a DOCX file into Tiptap JSON (runtime model).
 *
 * Combines @office-open/docx's `parseDocument` (DOCX binary → DocumentOptions)
 * with `DocxManager.resolve` (DocumentOptions → Tiptap JSON).
 */
export function parseDOCX(data: Parameters<typeof parseDocument>[0]): JSONContent {
  return defaultManager.resolve(parseDocument(data));
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
   * properties (`title`/`creator`/`description`/…), `styles`/`externalStyles`,
   * `background`, `features`, `fonts`, etc. Excludes `sections` (always compiled
   * from the JSON) and `numbering` (collected from ordered-list nodes).
   *
   * `styles`/`externalStyles` here take precedence over any `styles` carried on
   * `json.attrs.styles` (e.g. from a prior `parseDOCX`); the two are mutually
   * exclusive, so specifying `externalStyles` drops the compiled `styles`.
   */
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">;
}

/**
 * Merge {@link DocxGenerateOptions.document} into the `DocumentOptions`
 * compiled from Tiptap JSON.
 *
 * - `sections`/`numbering`: compile-owned (excluded from `document`, never
 *   overridden).
 * - `styles`/`externalStyles`: option wins over `json.attrs.styles`; the two
 *   are mutually exclusive, so `externalStyles` clears compiled `styles`.
 * - Everything else (core properties, background, features, …): injected.
 */
function applyDocumentOptions(
  base: DocumentOptions,
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">,
): DocumentOptions {
  if (!document) return base;
  const merged: DocumentOptions = { ...base, ...document };
  // externalStyles is mutually exclusive with styles — drop compiled styles.
  if (document.externalStyles !== undefined) {
    delete (merged as Partial<DocumentOptions>).styles;
  }
  return merged;
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
  const { prepare = true, packer, document } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocument(applyDocumentOptions(compileDocument(json), document), packer);
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
  const { packer, document } = options ?? {};
  return generateDocumentSync(applyDocumentOptions(compileDocument(json), document), packer);
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
  const { prepare = true, packer, document } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocumentStream(applyDocumentOptions(compileDocument(json), document), packer);
}

/**
 * Convert DocumentOptions (persistence model) to Tiptap JSON (runtime model).
 */
export function resolveDocument(docOpts: DocumentOptions): JSONContent {
  return defaultManager.resolve(docOpts);
}

/**
 * Convert Tiptap JSON (runtime model) to DocumentOptions (persistence model).
 */
export function compileDocument(json: JSONContent): DocumentOptions {
  return defaultManager.compile(json);
}
