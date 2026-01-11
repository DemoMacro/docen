import { TableCell } from "docx";
import { TableCellNode } from "../types";
import { convertParagraph } from "./paragraph";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table cell node to DOCX TableCell
 *
 * @param node - TipTap table cell node
 * @param options - Table options from PropertiesOptions
 * @returns DOCX TableCell object
 */
export function convertTableCell(
  node: TableCellNode,
  options: DocxExportOptions["table"],
): TableCell {
  // Convert paragraphs in the cell
  const paragraphs =
    node.content?.map((p) =>
      convertParagraph(
        p,
        options?.cell?.paragraph ??
          options?.row?.paragraph ??
          options?.paragraph,
      ),
    ) || [];

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
  // colwidth is an array of column widths, use the first one for this cell
  if (
    node.attrs?.colwidth !== null &&
    node.attrs?.colwidth !== undefined &&
    node.attrs.colwidth.length > 0
  ) {
    cellOptions.width = {
      size: node.attrs.colwidth[0],
      type: "dxa" as const,
    };
  }

  return new TableCell(cellOptions);
}
