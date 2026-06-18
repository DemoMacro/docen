import {
  generateDocument,
  generateDocumentStream,
  generateDocumentSync,
  parseDocument,
} from "@office-open/docx";
import type {
  DocumentOptions,
  SectionChild,
  ParagraphOptions,
  ParagraphChild,
  RunOptions,
  TableOptions,
  TableRowOptions,
  TableCellOptions,
  LevelsOptions,
  NumberingOptions,
  OutputByType,
  OutputType,
  PackerOptions,
} from "@office-open/docx";

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
import { bytesToBase64 } from "./base64";
import { prepareDocument, type PrepareStep } from "./prepare";

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

  compile(json: JSONContent): DocumentOptions {
    this.numberingConfigs = [];
    this.orderedInstanceCounter = 0;

    const children: SectionChild[] = [];

    if (json.content) {
      for (const node of json.content) {
        const child = this.compileSectionChild(node);
        if (!child) continue;
        if (Array.isArray(child)) children.push(...child);
        else children.push(child);
      }
    }

    return {
      sections: [{ children }],
      ...(this.numberingConfigs.length > 0
        ? { numbering: { config: this.numberingConfigs } as NumberingOptions }
        : {}),
    };
  }

  resolve(docOpts: DocumentOptions): JSONContent {
    const sections = docOpts.sections ?? [];
    if (sections.length === 0) {
      return { type: "doc", content: [{ type: "paragraph" }] };
    }

    const children = sections[0].children ?? [];
    const numberingLookup = this.buildNumberingLookup(docOpts);
    const content = this.resolveSectionChildren(children, numberingLookup);

    return {
      type: "doc",
      content: content.length > 0 ? content : [{ type: "paragraph" }],
    };
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
    const rows: Record<string, unknown>[] = [];

    // Track active vertical spans from previous rows
    let activeSpans: { colStart: number; colspan: number; remainingRows: number }[] = [];

    for (const rowNode of node.content ?? []) {
      if (rowNode.type !== "tableRow") continue;

      const rowOpts = tableRowExt.renderDocx(rowNode) as Record<string, unknown>;
      const pmCells = (rowNode.content ?? []).filter(
        (c) => c.type === "tableCell" || c.type === "tableHeader",
      );

      // Snapshot spans from previous rows for this row
      const currentSpans = [...activeSpans].sort((a, b) => a.colStart - b.colStart);
      const newSpans: { colStart: number; colspan: number; remainingRows: number }[] = [];
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
            children: [{ paragraph: "" }],
          });
          colIdx += span.colspan;
          spanIdx++;
        } else {
          // Compile and place the actual cell
          const cell = this.compileTableCellNode(pmCells[cellIdx]);
          const cs = (cell.columnSpan as number) ?? 1;
          const rs = (cell.rowSpan as number) ?? 1;

          if (rs > 1) {
            delete cell.rowSpan;
            cell.verticalMerge = "restart";
            newSpans.push({ colStart: colIdx, colspan: cs, remainingRows: rs - 1 });
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

  private compileTableCellNode(cellNode: JSONContent): Record<string, unknown> {
    const cellOpts = (
      cellNode.type === "tableHeader"
        ? tableHeaderExt.renderDocx(cellNode)
        : tableCellExt.renderDocx(cellNode)
    ) as Record<string, unknown>;

    // Cell horizontal alignment (Tiptap base-extension `align` attr) is NOT an
    // OOXML cell property — <w:tcPr> has no horizontal alignment. Push it down to
    // each contained paragraph's `alignment` (the OOXML <w:jc>), unless a paragraph
    // already specifies its own alignment.
    const cellAlign = (cellNode.attrs?.align as string | undefined) ?? null;

    const cellChildren: SectionChild[] = [];
    for (const paraNode of cellNode.content ?? []) {
      const para = this.compileParagraphNode(paraNode);
      const paraObj = (typeof para === "string" ? { text: para } : para) as Record<string, unknown>;
      if (cellAlign && !paraObj.alignment) paraObj.alignment = cellAlign;
      cellChildren.push({ paragraph: paraObj });
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
    let ordered: { reference: string; instance: number } | undefined;
    if (isOrdered) ordered = this.registerOrderedNumbering(node);

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

          if (ordered) {
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
        case "image": {
          const imageRun = imageExt.renderDocx(node);
          if (imageRun) children.push(imageRun);
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
    // rawXml (incl. aggregated TOC field), bookmarkStart/End, toc, textbox,
    // altChunk, subDoc, customXml — no native Tiptap node. Carry the SectionChild
    // verbatim so the round-trip is byte-faithful.
    return this.resolvePassthrough(child);
  }

  /** Wrap an opaque SectionChild in a passthrough atom (attrs.data = JSON). */
  private resolvePassthrough(child: SectionChild): JSONContent {
    return { type: "passthrough", attrs: { data: JSON.stringify(child) } };
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

    // Detect heading
    const headingLevel = resolved.heading ? HEADING_LEVEL_MAP[resolved.heading] : undefined;
    const nodeType = headingLevel ? "heading" : "paragraph";

    // Dispatch to extension parseDocx
    const attrs = headingLevel
      ? headingExt.parseDocx(resolved as unknown as Record<string, unknown>)
      : paragraphExt.parseDocx(resolved as unknown as Record<string, unknown>);

    // List paragraphs never reach here — resolveSectionChildren intercepts
    // them upstream and rebuilds the nested list tree.
    const content = this.resolveInlineContent(resolved);
    const cleanAttrsObj = cleanAttrs(attrs);

    const node: JSONContent = { type: nodeType };
    if (Object.keys(cleanAttrsObj).length > 0) node.attrs = cleanAttrsObj;
    if (content.length > 0) node.content = content;

    return node;
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
   * Walk section children, grouping consecutive list paragraphs into nested
   * Tiptap lists. Non-list children resolve individually. DOCX flattens lists
   * to a paragraph sequence (depth carried by `level`); this rebuilds the tree.
   */
  private resolveSectionChildren(
    children: SectionChild[],
    numberingLookup: Map<string, { format?: string; start?: number }>,
  ): JSONContent[] {
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
        firstPara && typeof firstPara !== "string"
          ? this.detectList(firstPara, numberingLookup)
          : null;

      if (!firstInfo) {
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
        const info = this.detectList(para, numberingLookup);
        if (!info) break;
        group.push({ para, info });
        i++;
      }
      content.push(...this.buildListTree(group));
    }
    return content;
  }

  /** Classify a paragraph as a list item, or null if it isn't one. */
  private detectList(
    para: ParagraphOptions,
    lookup: Map<string, { format?: string; start?: number }>,
  ): ListInfo | null {
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
      const cfg = reference ? lookup.get(reference) : undefined;
      // A config whose format isn't "bullet" → ordered; otherwise this is the
      // built-in default-bullet numbering (parse may tag numId=1 as numbering
      // when its abstractNum resolves), so degrade to bullet.
      if (cfg && cfg.format && cfg.format !== "bullet") {
        kind = "ordered";
        start = cfg.start;
      } else {
        kind = "bullet";
        reference = undefined;
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
        // Only level-0 ordered lists carry `start`; deeper levels restart at 1.
        if (
          listType === "orderedList" &&
          info.level === 0 &&
          typeof info.start === "number" &&
          info.start !== 1
        ) {
          newList.attrs = { start: info.start };
        }
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

  /**
   * Resolve a list-item paragraph to a Tiptap paragraph/heading node, stripping
   * the list marker (bullet/numbering) and the leading task checkbox — those
   * are expressed at the list/item level, not inside the paragraph.
   */
  private resolveListItemParagraph(para: ParagraphOptions, info: ListInfo): JSONContent {
    const resolved = typeof para === "string" ? ({ text: para } as ParagraphOptions) : para;
    const headingLevel = resolved.heading ? HEADING_LEVEL_MAP[resolved.heading] : undefined;
    const nodeType = headingLevel ? "heading" : "paragraph";

    const attrs = headingLevel
      ? headingExt.parseDocx(resolved as unknown as Record<string, unknown>)
      : paragraphExt.parseDocx(resolved as unknown as Record<string, unknown>);

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

    // DOCX encodes a row span as a `restart` cell followed by N empty `continue`
    // cells; ProseMirror uses one cell with rowspan = N+1. Track open spans per
    // column so each continuation cell increments the owning cell's `rowspan`,
    // which compileTableNode reads back to rebuild the continuation cells.
    let activeSpans = new Map<number, JSONContent>();

    for (const row of rows) {
      const rowAttrs = tableRowExt.parseDocx(row as Partial<TableRowOptions>);
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

        const cellChildren = (cell.children ?? []) as SectionChild[];
        const cellContent: JSONContent[] = [];
        for (const cc of cellChildren) {
          const resolved = this.resolveSectionChild(cc);
          if (resolved) cellContent.push(resolved);
        }

        const cellType = isHeader ? "tableHeader" : "tableCell";
        const cellNode: JSONContent = { type: cellType };
        if (Object.keys(cellAttrs).length > 0) cellNode.attrs = cleanAttrs(cellAttrs);
        if (cellContent.length > 0) cellNode.content = cellContent;

        if (vMerge === "restart") {
          // rowspan is finalized when continuation cells arrive below; register
          // the node so they can find and increment it.
          for (let c = colIdx; c < colIdx + cellColspan; c++) nextActiveSpans.set(c, cellNode);
        }

        cellNodes.push(cellNode);
        colIdx += cellColspan;
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
    if ("text" in child || "children" in child || "break" in child) {
      return this.resolveRun(child as RunOptions);
    }
    if ("image" in child) {
      return this.resolveImage(child.image as unknown as Record<string, unknown>);
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
    return null;
  }

  /** Resolve an inline SDT (mention carrier; other inline SDTs unsupported). */
  private resolveInlineSdt(child: ParagraphChild): JSONContent | null {
    if (mentionExt.isMention(child)) {
      const { id, label } = mentionExt.readMention(child);
      return { type: "mention", attrs: { id, label } };
    }
    return null;
  }

  private resolveRun(opts: RunOptions): JSONContent | null {
    // Pure break (no text/children) → hardBreak node
    if (opts.break && opts.text === undefined && !opts.children) {
      return { type: "hardBreak" };
    }
    const text = opts.text;
    if (text === undefined && !opts.children) return null;

    if (opts.children && !text) {
      const parts: string[] = [];
      for (const c of opts.children) {
        if (typeof c === "string") parts.push(c);
      }
      if (parts.length === 0) return null;
      return { type: "text", text: parts.join(""), marks: this.resolveMarks(opts) };
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

    if (opts.font === "Consolas" || opts.font === "monospace") {
      marks.push({ type: "code" });
    }

    const textStyleAttrs = textStyleExt.parseDocx(opts);
    // parseDocx passes RunStylePropertiesOptions through verbatim (size/color/font/
    // characterSpacing/rightToLeft/…). The `code` mark (Consolas/monospace) is
    // represented separately above — drop it from text-style attrs to avoid duplication.
    // color normalization happens in renderHTML via normalizeColorToHex.
    if (opts.font === "Consolas" || opts.font === "monospace") {
      delete textStyleAttrs.font;
    }

    if (Object.keys(textStyleAttrs).length > 0) {
      marks.push({ type: "textStyle", attrs: textStyleAttrs });
    }

    return marks.length > 0 ? marks : undefined;
  }

  private resolveImage(imageOpts: Record<string, unknown>): JSONContent {
    const attrs = imageExt.parseDocx(imageOpts);

    // Image data → data URL (chunked base64 — large images overflow the call
    // stack if spread into String.fromCharCode; see bytesToBase64).
    const data = imageOpts.data as Uint8Array | undefined;
    const type = imageOpts.type as string | undefined;
    if (data && type) {
      const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
      attrs.src = `data:image/${type};base64,${bytesToBase64(bytes)}`;
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
                target: "_blank",
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

const HEADING_LEVEL_MAP: Record<string, 1 | 2 | 3 | 4 | 5 | 6> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Title: 1,
};

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
  const { prepare = true, packer } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocument(compileDocument(json), packer);
}

/**
 * Generate a DOCX file synchronously — fastest throughput, blocks the event loop.
 *
 * Pipeline: `DocxManager.compile` → `generateDocumentSync`. Does **not** run
 * `prepareDocument` (it is async); call `await prepareDocument(json)` first
 * when http images need embedding.
 */
export function generateDOCXSync<T extends OutputType = "nodebuffer">(
  json: JSONContent,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return generateDocumentSync(compileDocument(json), packerOptions);
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
  const { prepare = true, packer } = options ?? {};
  if (prepare !== false) {
    await prepareDocument(json, prepare === true ? undefined : prepare);
  }
  return generateDocumentStream(compileDocument(json), packer);
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
