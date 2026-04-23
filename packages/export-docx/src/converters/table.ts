import { Table, TableRow, ITableOptions, AlignmentType, TableLayoutType } from "docx-plus";
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

  const rowErrors = rowResults.map((r, i) => ({ r, i })).filter(({ r }) => r.status === "rejected");
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

  // Apply table layout
  if (node.attrs?.layout) {
    const layoutMap: Record<string, (typeof TableLayoutType)[keyof typeof TableLayoutType]> = {
      autofit: TableLayoutType.AUTOFIT,
      fixed: TableLayoutType.FIXED,
    };
    const layoutVal = layoutMap[node.attrs.layout];
    if (layoutVal) {
      tableOptions = { ...tableOptions, layout: layoutVal };
    }
  }

  // Apply table alignment
  if (node.attrs?.alignment) {
    const alignMap: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
      left: AlignmentType.LEFT,
      center: AlignmentType.CENTER,
      right: AlignmentType.RIGHT,
    };
    const alignVal = alignMap[node.attrs.alignment];
    if (alignVal) {
      tableOptions = { ...tableOptions, alignment: alignVal };
    }
  }

  // Apply cell spacing
  if (node.attrs?.cellSpacing) {
    tableOptions = {
      ...tableOptions,
      cellSpacing: { value: node.attrs.cellSpacing, type: "dxa" },
    };
  }

  return new Table(tableOptions);
}
