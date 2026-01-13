import { TableCell } from "docx";
import { TableCellNode } from "../types";
import { convertParagraph } from "./paragraph";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table cell node to DOCX TableCell
 *
 * @param node - TipTap table cell node
 * @param options - Table options from PropertiesOptions
 * @param exportOptions - Export options (for image processing)
 * @returns Promise<DOCX TableCell object>
 */
export async function convertTableCell(
  node: TableCellNode,
  options: DocxExportOptions["table"],
  exportOptions?: DocxExportOptions,
): Promise<TableCell> {
  // Convert paragraphs in the cell
  const paragraphs = await Promise.all(
    (node.content || []).map((p) =>
      convertParagraph(
        p,
        options?.cell?.paragraph ?? options?.row?.paragraph ?? options?.paragraph,
        exportOptions,
      ),
    ),
  );

  // Create table cell options
  const cellOptions = {
    children: paragraphs,
    ...options?.cell?.run,
  };

  // Add column span if present
  if (node.attrs?.colspan && node.attrs.colspan > 1) {
    cellOptions.columnSpan = node.attrs.colspan;
  }

  // Add row span if present
  if (node.attrs?.rowspan && node.attrs.rowspan > 1) {
    cellOptions.rowSpan = node.attrs.rowspan;
  }

  // Add column width if present
  // colwidth can be a number (pixels) or an array of column widths
  if (node.attrs?.colwidth !== null && node.attrs?.colwidth !== undefined) {
    // Handle both number and array formats
    const widthInPixels = Array.isArray(node.attrs.colwidth)
      ? node.attrs.colwidth[0]
      : node.attrs.colwidth;

    if (widthInPixels && widthInPixels > 0) {
      // Convert pixels to twips (1 inch = 96 pixels = 1440 twips at 96 DPI)
      const twips = Math.round(widthInPixels * 15);
      cellOptions.width = {
        size: twips,
        type: "dxa" as const,
      };
    }
  }

  return new TableCell(cellOptions);
}
