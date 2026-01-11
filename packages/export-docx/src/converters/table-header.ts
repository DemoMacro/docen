import { TableCell } from "docx";
import { TableHeaderNode } from "../types";
import { convertParagraph } from "./paragraph";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table header node to DOCX TableCell
 *
 * @param node - TipTap table header node
 * @param options - Table options from PropertiesOptions
 * @returns DOCX TableCell object for header
 */
export function convertTableHeader(
  node: TableHeaderNode,
  options: DocxExportOptions["table"],
): TableCell {
  // Convert paragraphs in the header
  const paragraphs =
    node.content?.map((p) =>
      convertParagraph(
        p,
        options?.header?.paragraph ??
          options?.cell?.paragraph ??
          options?.row?.paragraph ??
          options?.paragraph,
      ),
    ) || [];

  // Create table header cell options
  const headerCellOptions = {
    children: paragraphs,
    ...options?.header?.run,
  };

  // Add column span if present
  if (node.attrs?.colspan && node.attrs.colspan > 1) {
    headerCellOptions.columnSpan = node.attrs.colspan;
  }

  // Add row span if present
  if (node.attrs?.rowspan && node.attrs.rowspan > 1) {
    headerCellOptions.rowSpan = node.attrs.rowspan;
  }

  // Add column width if present
  // colwidth is an array of column widths, use the first one for this cell
  if (
    node.attrs?.colwidth !== null &&
    node.attrs?.colwidth !== undefined &&
    node.attrs.colwidth.length > 0
  ) {
    headerCellOptions.width = {
      size: node.attrs.colwidth[0],
      type: "dxa" as const,
    };
  }

  return new TableCell(headerCellOptions);
}
