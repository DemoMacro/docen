import type { Element } from "xast";
import type { JSONContent } from "@tiptap/core";
import { convertParagraph } from "./paragraph";

/**
 * Check if an element is a table
 */
export function isTable(node: Element): boolean {
  return node.name === "w:tbl";
}

/**
 * Convert a table element to TipTap JSON
 */
export function convertTable(
  node: Element,
  hyperlinks: Map<string, string>,
  images: Map<string, string>,
): JSONContent {
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
  rowElements.forEach((rowElement, rowIndex) => {
    rows.push(
      convertTableRow(
        rowElement,
        rowIndex === 0,
        hyperlinks,
        images,
        activeRowspans,
        rowElements,
        rowIndex,
      ),
    );
  });

  return {
    type: "table",
    content: rows,
  };
}

/**
 * Convert a table row to TipTap JSON
 */
function convertTableRow(
  rowNode: Element,
  isFirstRow: boolean,
  hyperlinks: Map<string, string>,
  images: Map<string, string>,
  activeRowspans: Map<number, number>,
  allRows: Element[],
  currentRowIndex: number,
): JSONContent {
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
      const paragraphs = convertCellContent(child, hyperlinks, images);

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

  // Find cell properties element
  let tcPr: Element | undefined;
  for (const child of cellNode.children) {
    if (child.type === "element" && child.name === "w:tcPr") {
      tcPr = child;
      break;
    }
  }

  if (!tcPr) {
    return props;
  }

  // Check for gridSpan (colspan)
  for (const child of tcPr.children) {
    if (child.type === "element" && child.name === "w:gridSpan") {
      const val = child.attributes["w:val"];
      if (val) {
        props.colspan = parseInt(val as string);
      }
      break;
    }
  }

  // Check for vMerge (rowspan)
  // DOCX format: cells with vMerge/@val='continue' are merged into the cell above
  // Cells without vMerge or with vMerge but no val attribute are normal cells
  for (const child of tcPr.children) {
    if (child.type === "element" && child.name === "w:vMerge") {
      const val = child.attributes["w:val"];
      // If vMerge has val="continue", this cell is merged (rowspan should be 0 or handled specially)
      // If vMerge exists but no val attribute, or val is not "continue", it's a normal cell
      if (val === "continue") {
        props.rowspan = 0; // This cell is merged, should be skipped or handled specially
      }
      // Note: We don't set rowspan > 1 here because DOCX doesn't explicitly store the rowspan value
      // The rowspan value needs to be calculated by counting consecutive "continue" cells
      break;
    }
  }

  // Check for column width
  for (const child of tcPr.children) {
    if (child.type === "element" && child.name === "w:tcW") {
      const val = child.attributes["w:w"];
      if (val) {
        const twips = parseInt(val as string);
        // Convert twips to pixels (1 inch = 1440 twips = 96 pixels at 96 DPI)
        props.colwidth = Math.round(twips / 15);
      }
      break;
    }
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
function convertCellContent(
  cellNode: Element,
  hyperlinks: Map<string, string>,
  images: Map<string, string>,
): JSONContent[] {
  // Find all paragraphs in the cell
  const paragraphs: JSONContent[] = [];

  for (const child of cellNode.children) {
    if (child.type === "element" && child.name === "w:p") {
      const paragraph = convertParagraph(child, hyperlinks, images);
      paragraphs.push(paragraph);
    }
  }

  // Return all paragraphs to preserve complete cell content
  // DOCX cells can contain multiple paragraphs
  return paragraphs.length > 0 ? paragraphs : [{ type: "paragraph", content: [] }];
}
