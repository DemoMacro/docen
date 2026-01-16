import { TableRow } from "docx";
import { TableRowNode } from "../types";
import { convertTableCell } from "./table-cell";
import { convertTableHeader } from "./table-header";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table row node to DOCX TableRow
 *
 * @param node - TipTap table row node
 * @param params - Conversion parameters
 * @returns Promise<DOCX TableRow object>
 */
export async function convertTableRow(
  node: TableRowNode,
  params: {
    options: DocxExportOptions["table"];
  },
): Promise<TableRow> {
  const { options } = params;

  // Choose row options
  const rowOptions = options?.row;

  // Convert table cells and headers
  const cells = await Promise.all(
    (node.content || []).map(async (cellNode) => {
      if (cellNode.type === "tableCell") {
        return await convertTableCell(cellNode, params);
      } else if (cellNode.type === "tableHeader") {
        return await convertTableHeader(cellNode, params);
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
