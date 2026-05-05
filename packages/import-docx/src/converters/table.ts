import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { ParseContext } from "../parser";
import { convertParagraph } from "./paragraph";
import { parseTableProperties, parseRowProperties, parseCellProperties } from "../parsers/table";

/**
 * Check if an element is a table
 */
export function isTable(node: Element): boolean {
  return node.name === "w:tbl";
}

/**
 * Convert a table element to TipTap JSON
 */
export async function convertTable(
  node: Element,
  params: { context: ParseContext },
): Promise<JSONContent> {
  // Collect all rows first to enable rowspan calculation
  const rows: Element[] = [];
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:tr") {
      rows.push(child as Element);
    }
  }

  const activeRowspans = new Map<number, number>();

  // Process rows sequentially to ensure activeRowspans consistency.
  // activeRowspans tracks which columns are occupied by vertical merges from
  // previous rows. Concurrent processing (Promise.allSettled) causes race
  // conditions where a row reads stale activeRowspans values before earlier
  // rows have finished updating them.
  const content: JSONContent[] = [];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    try {
      const rowJson = await convertTableRow(rows[rowIndex], {
        context: params.context,
        activeRowspans,
        rows,
        rowIndex,
      });
      content.push(rowJson);
    } catch (reason) {
      throw new Error(`Failed to convert table rows:\n[row ${rowIndex}]: ${String(reason)}`);
    }
  }

  // Parse table properties (including cell margins)
  const tableProps = parseTableProperties(node);

  return {
    type: "table",
    ...(tableProps && { attrs: tableProps }),
    content,
  };
}

/**
 * Convert a table row to TipTap JSON
 */
async function convertTableRow(
  rowNode: Element,
  params: {
    context: ParseContext;
    activeRowspans: Map<number, number>;
    rows: Element[];
    rowIndex: number;
  },
): Promise<JSONContent> {
  const cells: JSONContent[] = [];
  let colIndex = 0;

  const rowProps = parseRowProperties(rowNode);

  for (const child of rowNode.children) {
    if (child.type !== "element" || child.name !== "w:tc") continue;

    const mergedBy = params.activeRowspans.get(colIndex);
    if (mergedBy && mergedBy > 0) {
      params.activeRowspans.set(colIndex, mergedBy - 1);
      colIndex++;
      continue;
    }

    let cellProps = parseCellProperties(child);

    if (cellProps?.rowspan === 1) {
      const actualRowSpan = calculateRowspan({
        rows: params.rows,
        rowIndex: params.rowIndex,
        colIndex,
      });
      if (actualRowSpan > 1) {
        cellProps = { ...cellProps, rowspan: actualRowSpan };
      }
    }

    if (cellProps?.rowspan && cellProps.rowspan > 1) {
      params.activeRowspans.set(colIndex, cellProps.rowspan - 1);
    }

    if (cellProps?.rowspan === 0) {
      colIndex++;
      continue;
    }

    const paragraphs = await convertCellContent(child, params);

    cells.push({
      type: "tableCell",
      ...(cellProps && { attrs: cellProps }),
      content: paragraphs,
    });

    colIndex += cellProps?.colspan || 1;
  }

  return {
    type: "tableRow",
    ...(rowProps && { attrs: rowProps }),
    content: cells,
  };
}

/**
 * Calculate the actual rowspan of a cell
 */
function calculateRowspan(params: { rows: Element[]; rowIndex: number; colIndex: number }): number {
  let rowspan = 1;
  let colIndex = params.colIndex;

  for (let rowIndex = params.rowIndex + 1; rowIndex < params.rows.length; rowIndex++) {
    const row = params.rows[rowIndex];
    let cellFound = false;
    let currentColIndex = colIndex; // Reset colIndex for each row

    for (const child of row.children) {
      if (child.type !== "element" || child.name !== "w:tc") continue;

      const cellProps = parseCellProperties(child);
      const colSpan = cellProps?.colspan || 1;

      if (currentColIndex >= 0 && currentColIndex < colSpan) {
        if (cellProps?.rowspan === 0) {
          rowspan++;
          cellFound = true;
        } else {
          return rowspan;
        }
        break;
      }

      currentColIndex -= colSpan;
    }

    if (!cellFound) break;
  }

  return rowspan;
}

/**
 * Convert cell content (typically paragraphs)
 */
async function convertCellContent(
  cellNode: Element,
  params: { context: ParseContext },
): Promise<JSONContent[]> {
  const paragraphs: JSONContent[] = [];

  for (const child of cellNode.children) {
    if (child.type === "element" && child.name === "w:p") {
      const paragraph = await convertParagraph(child, params);
      if (Array.isArray(paragraph)) {
        paragraphs.push(...paragraph);
      } else {
        paragraphs.push(paragraph);
      }
    }
  }

  return paragraphs.length ? paragraphs : [{ type: "paragraph", content: [] }];
}
