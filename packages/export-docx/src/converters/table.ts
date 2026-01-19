import { Table, ITableOptions } from "docx";
import { TableNode } from "@docen/extensions/types";
import { convertTableRow } from "./table-row";
import { DocxExportOptions } from "../options";

/**
 * Apply table margins to table options
 */

const applyTableMargins = <T extends ITableOptions>(options: T, node: TableNode): T => {
  const margins = {
    top: node.attrs?.marginTop ?? undefined,
    bottom: node.attrs?.marginBottom ?? undefined,
    left: node.attrs?.marginLeft ?? undefined,
    right: node.attrs?.marginRight ?? undefined,
  };

  // Only add margins if at least one is defined
  if (margins.top || margins.bottom || margins.left || margins.right) {
    return {
      ...options,
      margins,
    };
  }

  return options;
};

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
  let tableOptions: ITableOptions = {
    rows,
    ...options?.run,
  };

  // Apply table margins
  tableOptions = applyTableMargins(tableOptions, node);

  return new Table(tableOptions);
}
