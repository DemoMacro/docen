import { TableRow, TableCell, type ITableRowOptions } from "docx-plus";
import { TableRowNode } from "@docen/extensions/types";
import { convertTableCell } from "./table-cell";
import { convertTableHeader } from "./table-header";
import { DocxExportOptions } from "../options";
import { convertCssLengthToPixels, convertPixelsToTwip } from "../utils";

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
  const cellResults = await Promise.allSettled(
    (node.content || []).map(async (cellNode) => {
      if (cellNode.type === "tableCell") {
        return await convertTableCell(cellNode, params);
      } else if (cellNode.type === "tableHeader") {
        return await convertTableHeader(cellNode, params);
      }
      return null;
    }),
  );

  const cellErrors = cellResults
    .map((r, i) => ({ r, i, type: node.content?.[i]?.type }))
    .filter(({ r }) => r.status === "rejected");
  if (cellErrors.length > 0) {
    const msgs = cellErrors.map(
      ({ i, type, r }) => `[cell ${i}, type=${type}]: ${(r as PromiseRejectedResult).reason}`,
    );
    throw new Error(`Failed to convert table row cells:\n${msgs.join("\n")}`);
  }

  const cells = cellResults.map((r) => (r as PromiseFulfilledResult<TableCell | null>).value);

  // Filter out null values
  const validCells = cells.filter((cell): cell is TableCell => cell !== undefined);

  // Prepare table row options
  const height = (() => {
    if (!node.attrs?.rowHeight) return undefined;
    const twips = convertPixelsToTwip(convertCssLengthToPixels(node.attrs.rowHeight));
    if (twips <= 0) return undefined;
    const rule = node.attrs.rowHeightRule === "exact" ? "exact" as const : "atLeast" as const;
    return { rule, value: twips };
  })();

  const tableRowOptions: ITableRowOptions = {
    children: validCells,
    ...rowOptions,
    ...(height && { height }),
    ...(node.attrs?.header && { tableHeader: true }),
  };

  const row = new TableRow(tableRowOptions);

  return row;
}
