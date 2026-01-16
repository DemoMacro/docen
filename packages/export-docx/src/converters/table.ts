import { Table, ITableOptions } from "docx";
import { TableNode } from "../types";
import { convertTableRow } from "./table-row";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table node to DOCX Table
 *
 * @param node - TipTap table node
 * @param params - Conversion parameters
 * @returns Promise<Table>
 */
export async function convertTable(
  node: TableNode,
  params: {
    options: DocxExportOptions["table"];
  },
): Promise<Table> {
  const { options } = params;

  // Convert table rows
  const rows = await Promise.all((node.content || []).map((row) => convertTableRow(row, params)));

  // Build table options with options
  const tableOptions: ITableOptions = {
    rows,
    ...options?.run, // Apply table options
  };

  // Create and return table
  return new Table(tableOptions);
}
