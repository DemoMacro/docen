import { Table, Paragraph, ITableOptions } from "docx";
import { TableNode } from "../types";
import { convertTableRow } from "./table-row";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table node to DOCX Table
 *
 * @param node - TipTap table node
 * @param options - Table options from PropertiesOptions
 * @param exportOptions - Export options (for image processing)
 * @returns Promise<Array containing Table and a following Paragraph to prevent merging>
 */
export async function convertTable(
  node: TableNode,
  options: DocxExportOptions["table"],
  exportOptions?: DocxExportOptions,
): Promise<Array<Table | Paragraph>> {
  // Convert table rows
  const rows = await Promise.all(
    (node.content || []).map((row) => convertTableRow(row, options, exportOptions)),
  );

  // Build table options with options
  const tableOptions: ITableOptions = {
    rows,
    ...options?.run, // Apply table options
  };

  // Create table
  const table = new Table(tableOptions);

  // Return table with a following empty paragraph to prevent automatic merging with adjacent tables
  return [table, new Paragraph({})];
}
