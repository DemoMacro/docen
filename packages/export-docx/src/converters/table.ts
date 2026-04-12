import { Table, TableRow, ITableOptions } from "docx";
import { TableNode } from "@docen/extensions/types";
import { convertTableRow } from "./table-row";
import { DocxExportOptions } from "../options";

/**
 * Apply table margins to table options
 */

export const applyTableMargins = <T extends ITableOptions>(options: T, node: TableNode): T => {
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
  const rowResults = await Promise.allSettled(
    (node.content || []).map((row) => convertTableRow(row, params)),
  );

  const rowErrors = rowResults
    .map((r, i) => ({ r, i }))
    .filter(({ r }) => r.status === "rejected");
  if (rowErrors.length > 0) {
    const msgs = rowErrors.map(({ i, r }) => `[row ${i}]: ${(r as PromiseRejectedResult).reason}`);
    throw new Error(`Failed to convert table rows:\n${msgs.join("\n")}`);
  }

  const rows = rowResults.map((r) => (r as PromiseFulfilledResult<TableRow>).value);

  // Build table options
  let tableOptions: ITableOptions = {
    rows,
    // Apply table style if configured
    ...(options?.style?.id && { style: options.style.id }),
    ...options?.run,
  };

  // Apply table margins
  tableOptions = applyTableMargins(tableOptions, node);

  return new Table(tableOptions);
}
