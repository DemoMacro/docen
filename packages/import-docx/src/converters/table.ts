import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "../options";
import type { StyleMap } from "../parsers/styles";
import type { ImageInfo } from "../parsers/types";
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
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, ImageInfo>;
    options?: DocxImportOptions;
    styleMap?: StyleMap;
  },
): Promise<JSONContent> {
  // Collect all rows first to enable rowspan calculation
  const rows: Element[] = [];
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:tr") {
      rows.push(child as Element);
    }
  }

  const activeRowspans = new Map<number, number>();

  // Convert each row
  const content = await Promise.all(
    rows.map((row, rowIndex) =>
      convertTableRow(row, {
        ...params,
        activeRowspans,
        rows,
        rowIndex,
      }),
    ),
  );

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
    hyperlinks: Map<string, string>;
    images: Map<string, ImageInfo>;
    options?: DocxImportOptions;
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

    if (cellProps?.rowSpan === 1) {
      const actualRowSpan = calculateRowspan({
        ...params,
        colIndex,
      });
      if (actualRowSpan > 1) {
        cellProps = { ...cellProps, rowSpan: actualRowSpan };
      }
    }

    if (cellProps?.rowSpan && cellProps.rowSpan > 1) {
      params.activeRowspans.set(colIndex, cellProps.rowSpan - 1);
    }

    if (cellProps?.rowSpan === 0) {
      colIndex++;
      continue;
    }

    const paragraphs = await convertCellContent(child, params);

    cells.push({
      type: "tableCell",
      ...(cellProps && { attrs: cellProps }),
      content: paragraphs,
    });

    colIndex += cellProps?.colSpan || 1;
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

    for (const child of row.children) {
      if (child.type !== "element" || child.name !== "w:tc") continue;

      const cellProps = parseCellProperties(child);
      const colSpan = cellProps?.colSpan || 1;

      if (colIndex >= 0 && colIndex < colSpan) {
        if (cellProps?.rowSpan === 0) {
          rowspan++;
          cellFound = true;
        } else {
          return rowspan;
        }
        break;
      }

      colIndex -= colSpan;
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
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, ImageInfo>;
    options?: DocxImportOptions;
  },
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
