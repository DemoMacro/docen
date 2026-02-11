import { TableCell, Paragraph, IParagraphOptions } from "docx";
import { TableCellNode } from "@docen/extensions/types";
import { convertParagraph } from "./paragraph";
import { convertBorder } from "../utils";
import { DocxExportOptions } from "../options";

/**
 * Convert TipTap table cell node to DOCX TableCell
 *
 * @param node - TipTap table cell node
 * @param params - Conversion parameters
 * @returns Promise<DOCX TableCell object>
 */
export async function convertTableCell(
  node: TableCellNode,
  params: {
    options: DocxExportOptions["table"];
  },
): Promise<TableCell> {
  const { options } = params;

  // Prepare paragraph options for table cells
  let cellParagraphOptions: IParagraphOptions =
    options?.cell?.paragraph ?? options?.row?.paragraph ?? {};

  // Apply style reference if configured
  if (options?.style) {
    cellParagraphOptions = {
      ...cellParagraphOptions,
      style: options.style.id,
    };
  }

  // Convert paragraphs in the cell
  const paragraphOptionsList = await Promise.all(
    (node.content || []).map((p) =>
      convertParagraph(p, {
        options: cellParagraphOptions,
      }),
    ),
  );

  // Convert IParagraphOptions[] to Paragraph[] for TableCell children
  const paragraphs = paragraphOptionsList.map((options) => new Paragraph(options));

  // Create table cell options
  const cellOptions = {
    children: paragraphs,
    ...options?.cell?.run,
  };

  // Add column span if present
  if (node.attrs?.colspan && node.attrs.colspan > 1) {
    cellOptions.columnSpan = node.attrs.colspan;
  }

  // Add row span if present
  if (node.attrs?.rowspan && node.attrs.rowspan > 1) {
    cellOptions.rowSpan = node.attrs.rowspan;
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
      cellOptions.width = {
        size: twips,
        type: "dxa" as const,
      };
    }
  }

  // Add background color if present
  if (node.attrs?.backgroundColor) {
    const hexColor = node.attrs.backgroundColor.replace("#", "");
    cellOptions.shading = { fill: hexColor };
  }

  // Add vertical alignment if present
  if (node.attrs?.verticalAlign) {
    // CSS "middle" → DOCX "center"
    const align = node.attrs.verticalAlign === "middle" ? "center" : node.attrs.verticalAlign;
    cellOptions.verticalAlign = align;
  }

  // Add borders if present
  const borders = {
    top: convertBorder(node.attrs?.borderTop),
    bottom: convertBorder(node.attrs?.borderBottom),
    left: convertBorder(node.attrs?.borderLeft),
    right: convertBorder(node.attrs?.borderRight),
  };

  if (borders.top || borders.bottom || borders.left || borders.right) {
    cellOptions.borders = borders;
  }

  return new TableCell(cellOptions);
}
