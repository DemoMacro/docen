import { TableCell } from "docx";
import { TableCellNode } from "../types";
import { convertParagraph } from "./paragraph";
import { PropertiesOptions } from "../option";

/**
 * Convert TipTap table cell node to DOCX TableCell
 *
 * @param node - TipTap table cell node
 * @param options - Table options from PropertiesOptions
 * @returns DOCX TableCell object
 */
export function convertTableCell(
  node: TableCellNode,
  options: PropertiesOptions["table"],
): TableCell {
  // Convert paragraphs in the cell
  const paragraphs = node.content?.map((p) => convertParagraph(p)) || [];

  // Create table cell with options
  const cell = new TableCell({
    children: paragraphs,
    ...options?.cell,
  });

  // Add column span if present
  if (node.attrs?.colspan && node.attrs.colspan > 1) {
    Object.assign(cell.options, { columnSpan: node.attrs.colspan });
  }

  // Add row span if present
  if (node.attrs?.rowspan && node.attrs.rowspan > 1) {
    Object.assign(cell.options, { rowSpan: node.attrs.rowspan });
  }

  // Add column width if present
  if (node.attrs?.colwidth !== null && node.attrs?.colwidth !== undefined) {
    Object.assign(cell.options, {
      width: {
        size: node.attrs.colwidth,
        type: "dxa" as const,
      },
    });
  }

  return cell;
}
