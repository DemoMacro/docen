import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import type { DocxImportOptions } from "../option";
import type { StyleMap } from "../parser";
import { convertParagraph } from "./paragraph";
import { findChild } from "../utils/xml";

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
    images: Map<string, string>;
    options?: DocxImportOptions;
    styleMap?: StyleMap;
  },
): Promise<JSONContent> {
  const rows: JSONContent[] = [];

  // Collect all rows first to enable rowspan calculation
  const rowElements: Element[] = [];
  for (const child of node.children) {
    if (child.type === "element" && child.name === "w:tr") {
      rowElements.push(child as Element);
    }
  }

  const activeRowspans = new Map<number, number>();

  // Convert each row
  for (let rowIndex = 0; rowIndex < rowElements.length; rowIndex++) {
    const rowElement = rowElements[rowIndex];
    rows.push(
      await convertTableRow(
        rowElement,
        rowIndex === 0,
        params,
        activeRowspans,
        rowElements,
        rowIndex,
      ),
    );
  }

  return {
    type: "table",
    content: rows,
  };
}

/**
 * Convert a table row to TipTap JSON
 */
async function convertTableRow(
  rowNode: Element,
  isFirstRow: boolean,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
  activeRowspans: Map<number, number>,
  allRows: Element[],
  currentRowIndex: number,
): Promise<JSONContent> {
  const cells: JSONContent[] = [];
  let colIndex = 0;

  // Find all table cells (w:tc)
  for (const child of rowNode.children) {
    if (child.type === "element" && child.name === "w:tc") {
      // Check if this cell is merged due to rowspan from above
      const mergedBy = activeRowspans.get(colIndex);
      if (mergedBy && mergedBy > 0) {
        // This cell is part of a rowspan from above, decrement and skip
        activeRowspans.set(colIndex, mergedBy - 1);
        colIndex++;
        continue;
      }

      // Get cell properties
      let cellProps = getCellProperties(child);

      // If this cell doesn't have vMerge (not a merged cell), calculate actual rowspan
      if (cellProps && cellProps.rowspan === 1) {
        const actualRowspan = calculateRowspan(allRows, currentRowIndex, colIndex);
        if (actualRowspan > 1) {
          cellProps = { ...cellProps, rowspan: actualRowspan };
        }
      }

      // Track rowspan for subsequent rows
      if (cellProps && cellProps.rowspan > 1) {
        activeRowspans.set(colIndex, cellProps.rowspan - 1);
      }

      // For merged cells (rowspan=0), they should already be skipped above
      // But if we somehow get here, skip them
      if (cellProps && cellProps.rowspan === 0) {
        colIndex++;
        continue;
      }

      // DOCX format doesn't distinguish between tableHeader and tableCell
      // All cells should be tableCell for accurate round-trip conversion
      const cellType = "tableCell";

      // Convert cell content
      const paragraphs = await convertCellContent(child, params);

      cells.push({
        type: cellType,
        ...(cellProps && { attrs: cellProps }),
        content: paragraphs,
      });

      // Move column index by colspan
      colIndex += cellProps?.colspan || 1;
    }
  }

  return {
    type: "tableRow",
    content: cells,
  };
}

/**
 * Get cell properties (colspan, rowspan, colwidth)
 */
function getCellProperties(cellNode: Element): {
  colspan: number;
  rowspan: number;
  colwidth: number | null;
} | null {
  const props = {
    colspan: 1,
    rowspan: 1,
    colwidth: null as number | null,
  };

  const tcPr = findChild(cellNode, "w:tcPr");
  if (!tcPr) return props;

  // Check for gridSpan (colspan)
  const gridSpan = findChild(tcPr, "w:gridSpan");
  if (gridSpan?.attributes["w:val"]) {
    props.colspan = parseInt(gridSpan.attributes["w:val"] as string);
  }

  // Check for vMerge (rowspan)
  // DOCX format: cells with vMerge/@val='continue' are merged into the cell above
  // Cells without vMerge or with vMerge but no val attribute are normal cells
  const vMerge = findChild(tcPr, "w:vMerge");
  if (vMerge?.attributes["w:val"] === "continue") {
    props.rowspan = 0; // This cell is merged, should be skipped or handled specially
  }
  // Note: We don't set rowspan > 1 here because DOCX doesn't explicitly store the rowspan value
  // The rowspan value needs to be calculated by counting consecutive "continue" cells

  // Check for column width
  const tcW = findChild(tcPr, "w:tcW");
  if (tcW?.attributes["w:w"]) {
    const twips = parseInt(tcW.attributes["w:w"] as string);
    // Convert twips to pixels (1 inch = 1440 twips = 96 pixels at 96 DPI)
    props.colwidth = Math.round(twips / 15);
  }

  return props;
}

/**
 * Calculate the actual rowspan of a cell by counting consecutive vMerge/continue cells below it
 */
function calculateRowspan(allRows: Element[], startRowIndex: number, colIndex: number): number {
  let rowspan = 1;
  let currentColIndex = colIndex;

  // Check each subsequent row
  for (let rowIndex = startRowIndex + 1; rowIndex < allRows.length; rowIndex++) {
    const row = allRows[rowIndex];
    let cellFound = false;

    // Find the cell at currentColIndex in this row
    for (const child of row.children) {
      if (child.type === "element" && child.name === "w:tc") {
        const cellProps = getCellProperties(child);
        const colspan = cellProps?.colspan || 1;

        // If this is the cell at currentColIndex
        if (currentColIndex >= 0 && currentColIndex < colspan) {
          // Check if this cell has vMerge/continue
          if (cellProps?.rowspan === 0) {
            rowspan++;
            cellFound = true;
          } else {
            // Not a merged cell, stop counting
            return rowspan;
          }
          break;
        }

        currentColIndex -= colspan;
      }
    }

    if (!cellFound) {
      break;
    }
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
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
): Promise<JSONContent[]> {
  // Find all paragraphs in the cell
  const paragraphs: JSONContent[] = [];

  for (const child of cellNode.children) {
    if (child.type === "element" && child.name === "w:p") {
      const paragraph = await convertParagraph(child, params);
      paragraphs.push(paragraph);
    }
  }

  // Return all paragraphs to preserve complete cell content
  // DOCX cells can contain multiple paragraphs
  return paragraphs.length > 0 ? paragraphs : [{ type: "paragraph", content: [] }];
}
