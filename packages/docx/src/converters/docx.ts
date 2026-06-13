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
  OutputByType,
  OutputType,
  PackerOptions,
} from "@office-open/docx";

import type { JSONContent } from "../core";
import * as headingExt from "../extensions/heading";
import * as imageExt from "../extensions/image";
// Node renderDocx/parseDocx
import * as paragraphExt from "../extensions/paragraph";
// Mark renderDocx/parseDocx
import * as strikeExt from "../extensions/strike";
import * as tableExt from "../extensions/table";
import * as tableCellExt from "../extensions/table-cell";
import * as tableHeaderExt from "../extensions/table-header";
import * as tableRowExt from "../extensions/table-row";
import * as textStyleExt from "../extensions/text-style";
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

// ── DocxManager ──

/**
 * Manages DOCX serialization (Tiptap JSON ↔ DocumentOptions).
 *
 * Each extension provides renderDocx/parseDocx for its own attrs mapping.
 * DocxManager handles tree walking, child assembly, and dispatching.
 */
export class DocxManager {
  compile(json: JSONContent): DocumentOptions {
    const children: SectionChild[] = [];

    if (json.content) {
      for (const node of json.content) {
        const child = this.compileSectionChild(node);
        if (!child) continue;
        if (Array.isArray(child)) children.push(...child);
        else children.push(child);
      }
    }

    return { sections: [{ children }] };
  }

  resolve(docOpts: DocumentOptions): JSONContent {
    const sections = docOpts.sections ?? [];
    if (sections.length === 0) {
      return { type: "doc", content: [{ type: "paragraph" }] };
    }

    const children = sections[0].children ?? [];
    const content: JSONContent[] = [];

    for (const child of children) {
      const node = this.resolveSectionChild(child);
      if (node) content.push(node);
    }

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
        // Flatten blockquote into paragraphs
        for (const child of node.content ?? []) {
          if (child.type === "paragraph") {
            return { paragraph: this.compileParagraphNode(child) };
          }
        }
        return null;
      }
      case "codeBlock":
        return { paragraph: this.compileCodeBlock(node) };
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
        return this.compileListFromNode(node);
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

  private compileCodeBlock(node: JSONContent): ParagraphOptions {
    const text = (node.content ?? []).map((c) => c.text ?? "").join("");
    const language = node.attrs?.language as string | undefined;
    return {
      style: "Code",
      ...(language ? { run: { font: "Consolas" } } : {}),
      text,
    };
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

  private compileListFromNode(node: JSONContent): SectionChild[] | null {
    const items: SectionChild[] = [];
    for (const listItem of node.content ?? []) {
      if (listItem.type !== "listItem" && listItem.type !== "taskItem") continue;
      for (const child of listItem.content ?? []) {
        if (child.type === "paragraph" || child.type === "heading") {
          const para = this.compileParagraphNode(child);
          // simplifyParagraph may return a bare string — wrap to allow property assignment
          const paraObj =
            typeof para === "string"
              ? ({ text: para } as Record<string, unknown>)
              : (para as Record<string, unknown>);
          if (node.type === "bulletList") {
            paraObj.bullet = { level: 0 };
          } else if (node.type === "orderedList") {
            paraObj.numbering = { reference: "default-numbering", level: 0 };
          }
          items.push({ paragraph: paraObj as ParagraphOptions });
        }
      }
    }
    return items.length > 0 ? items : null;
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
        case "image": {
          const imageRun = imageExt.renderDocx(node);
          if (imageRun) children.push(imageRun);
          break;
        }
      }
    }

    return children;
  }

  private compileTextNode(node: JSONContent, children: ParagraphChild[]): void {
    const text = node.text ?? "";
    if (!text) return;

    const marks = node.marks ?? [];
    const runOpts: Record<string, unknown> = { text };

    for (const mark of marks) {
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
    return null;
  }

  private resolveParagraph(opts: string | ParagraphOptions): JSONContent {
    const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;

    // Detect heading
    const headingLevel = resolved.heading ? HEADING_LEVEL_MAP[resolved.heading] : undefined;
    const nodeType = headingLevel ? "heading" : "paragraph";

    // Dispatch to extension parseDocx
    const attrs = headingLevel
      ? headingExt.parseDocx(resolved as unknown as Record<string, unknown>)
      : paragraphExt.parseDocx(resolved as unknown as Record<string, unknown>);

    // Handle list items
    if (resolved.bullet) {
      return this.resolveListItem(resolved, "bulletList", attrs);
    }
    if (resolved.numbering) {
      return this.resolveListItem(resolved, "orderedList", attrs);
    }

    // Resolve inline content (recovers paragraph-level `text` when collapsed)
    const content = this.resolveInlineContent(resolved);
    const cleanAttrsObj = cleanAttrs(attrs);

    const node: JSONContent = { type: nodeType };
    if (Object.keys(cleanAttrsObj).length > 0) node.attrs = cleanAttrsObj;
    if (content.length > 0) node.content = content;

    return node;
  }

  private resolveListItem(
    opts: ParagraphOptions,
    listType: "bulletList" | "orderedList",
    paraAttrs: Record<string, unknown>,
  ): JSONContent {
    const content = this.resolveInlineContent(opts);
    const cleanAttrsObj = cleanAttrs(paraAttrs);

    const paragraphNode: JSONContent = { type: "paragraph" };
    if (Object.keys(cleanAttrsObj).length > 0) paragraphNode.attrs = cleanAttrsObj;
    if (content.length > 0) paragraphNode.content = content;

    return {
      type: listType,
      content: [{ type: "listItem", content: [paragraphNode] }],
    };
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
      return { type: "paragraph", attrs: { pageBreak: true } };
    }
    if ("columnBreak" in child) {
      return { type: "hardBreak" };
    }
    return null;
  }

  private resolveRun(opts: RunOptions): JSONContent | null {
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
    if (opts.highlight) marks.push({ type: "highlight" });

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

    // Image data → data URL
    const data = imageOpts.data as Uint8Array | undefined;
    const type = imageOpts.type as string | undefined;
    if (data && type) {
      const base64 =
        typeof btoa !== "undefined"
          ? btoa(
              String.fromCharCode(...(data instanceof ArrayBuffer ? new Uint8Array(data) : data)),
            )
          : Buffer.from(data).toString("base64");
      attrs.src = `data:image/${type};base64,${base64}`;
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
