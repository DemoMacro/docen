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

  // Build table options
  const tableOptions: ITableOptions = {
    rows,
    ...options?.run,
  };

  // Add table cell margins if present
  if (
    node.attrs?.marginTop !== undefined ||
    node.attrs?.marginBottom !== undefined ||
    node.attrs?.marginLeft !== undefined ||
    node.attrs?.marginRight !== undefined
  ) {
    return new Table({
      ...tableOptions,
      margins: {
        top: node.attrs.marginTop ?? undefined,
        bottom: node.attrs.marginBottom ?? undefined,
        left: node.attrs.marginLeft ?? undefined,
        right: node.attrs.marginRight ?? undefined,
      },
    });
  }

  return new Table(tableOptions);
}
