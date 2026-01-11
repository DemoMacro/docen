import { TableRow } from "docx";
import { TableRowNode } from "../types";
import { convertTableCell } from "./table-cell";
import { convertTableHeader } from "./table-header";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table row node to DOCX TableRow
 *
 * @param node - TipTap table row node
 * @param options - Table options from PropertiesOptions
 * @param exportOptions - Export options (for image processing)
 * @returns Promise<DOCX TableRow object>
 */
export async function convertTableRow(
  node: TableRowNode,
  options: DocxExportOptions["table"],
  exportOptions?: DocxExportOptions,
): Promise<TableRow> {
  // Choose row options
  const rowOptions = options?.row;

  // Convert table cells and headers
  const cells = await Promise.all(
    (node.content || []).map(async (cellNode) => {
      if (cellNode.type === "tableCell") {
        return await convertTableCell(cellNode, options, exportOptions);
      } else if (cellNode.type === "tableHeader") {
        return await convertTableHeader(cellNode, options, exportOptions);
      }
      return null;
    }),
  );

  // Filter out null values
  const validCells = cells.filter((cell) => cell !== null);

  // Create table row with options
  const row = new TableRow({
    children: validCells,
    ...rowOptions,
  });

  return row;
}
