import { TableCell, Paragraph, IParagraphOptions } from "docx";
import { TableHeaderNode } from "@docen/extensions/types";
import { convertParagraph } from "./paragraph";
import { convertBorder } from "../utils";
import { DocxExportOptions } from "../options";

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
  const paragraphOptionsList = await Promise.all(
    (node.content || []).map((p) =>
      convertParagraph(p, {
        options: headerParagraphOptions,
      }),
    ),
  );

  // Convert IParagraphOptions[] to Paragraph[] for TableCell children
  const paragraphs = paragraphOptionsList.map((options) => new Paragraph(options));

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
  // colwidth is an array of column widths (TipTap standard)
  if (node.attrs?.colwidth !== null && node.attrs?.colwidth !== undefined) {
    // Handle array format - take first width for the cell
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

  // Add background color if present
  if (node.attrs?.backgroundColor) {
    const hexColor = node.attrs.backgroundColor.replace("#", "");
    headerCellOptions.shading = { fill: hexColor };
  }

  // Add vertical alignment if present
  if (node.attrs?.verticalAlign) {
    // CSS "middle" → DOCX "center"
    const align = node.attrs.verticalAlign === "middle" ? "center" : node.attrs.verticalAlign;
    headerCellOptions.verticalAlign = align;
  }

  // Add borders if present
  const borders: Record<string, { color?: string; size?: number; style?: string; space?: number }> = {};

  const top = convertBorder(node.attrs?.borderTop);
  if (top) borders.top = top;

  const bottom = convertBorder(node.attrs?.borderBottom);
  if (bottom) borders.bottom = bottom;

  const left = convertBorder(node.attrs?.borderLeft);
  if (left) borders.left = left;

  const right = convertBorder(node.attrs?.borderRight);
  if (right) borders.right = right;

  if (Object.keys(borders).length > 0) {
    headerCellOptions.borders = borders;
  }

  return new TableCell(headerCellOptions);
}
