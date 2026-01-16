import { TableCell, IParagraphOptions } from "docx";
import { TableHeaderNode } from "../types";
import { convertParagraph } from "./paragraph";
import { DocxExportOptions } from "../option";

/**
 * Convert TipTap table header node to DOCX TableCell
 *
 * @param node - TipTap table header node
 * @param params - Conversion parameters
 * @returns Promise<DOCX TableCell object for header>
 */
export async function convertTableHeader(
  node: TableHeaderNode,
  params: {
    options: DocxExportOptions["table"];
  },
): Promise<TableCell> {
  const { options } = params;

  // Prepare paragraph options for table header cells
  let headerParagraphOptions: IParagraphOptions =
    options?.header?.paragraph ?? options?.cell?.paragraph ?? options?.row?.paragraph ?? {};

  // Apply style reference if configured
  if (options?.style) {
    headerParagraphOptions = {
      ...headerParagraphOptions,
      style: options.style.id,
    };
  }

  // Convert paragraphs in the header
  const paragraphs = await Promise.all(
    (node.content || []).map((p) =>
      convertParagraph(p, {
        options: headerParagraphOptions,
      }),
    ),
  );

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
  // colwidth can be a number (pixels) or an array of column widths
  if (node.attrs?.colwidth !== null && node.attrs?.colwidth !== undefined) {
    // Handle both number and array formats
    const widthInPixels = Array.isArray(node.attrs.colwidth)
      ? node.attrs.colwidth[0]
      : node.attrs.colwidth;

    if (widthInPixels && widthInPixels > 0) {
      // Convert pixels to twips (1 inch = 96 pixels = 1440 twips at 96 DPI)
      const twips = Math.round(widthInPixels * 15);
      headerCellOptions.width = {
        size: twips,
        type: "dxa" as const,
      };
    }
  }

  return new TableCell(headerCellOptions);
}
