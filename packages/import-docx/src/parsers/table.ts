import type { Element } from "xast";
import type { TableCellBorder } from "@docen/extensions/types";
import { findChild } from "../utils/xml";

/**
 * Parse a single border element
 */
export function parseBorder(borderNode: Element | undefined): {
  color?: string;
  width?: number;
  style?: "solid" | "dashed" | "dotted" | "double" | "none";
} | null {
  if (!borderNode) return null;

  const val = borderNode.attributes["w:val"] as string;
  const size = borderNode.attributes["w:sz"] as string;
  const color = borderNode.attributes["w:color"] as string;

  // Map DOCX border styles to CSS border styles
  const styleMap: Record<string, "solid" | "dashed" | "dotted" | "double" | "none"> = {
    single: "solid",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    none: "none",
    nil: "none",
  };

  const border: {
    color?: string;
    width?: number;
    style?: "solid" | "dashed" | "dotted" | "double" | "none";
  } = {};

  if (color && color !== "auto") {
    border.color = `#${color}`;
  }

  if (size) {
    // DOCX size is in eighth-points
    // Convert to pixels: 1 eighth-point = 1/8 pt = 1/8 * (4/3) px = 1/6 px ≈ 0.167 px
    const eighthPoints = parseInt(size);
    if (!isNaN(eighthPoints)) {
      border.width = Math.round(eighthPoints / 6);
    }
  }

  if (val && styleMap[val]) {
    border.style = styleMap[val];
  }

  return Object.keys(border).length > 0 ? border : null;
}

/**
 * Get table properties (cell margins)
 */
export function parseTableProperties(tableNode: Element): {
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
} | null {
  const props = {
    marginTop: undefined as number | undefined,
    marginBottom: undefined as number | undefined,
    marginLeft: undefined as number | undefined,
    marginRight: undefined as number | undefined,
  };

  const tblPr = findChild(tableNode, "w:tblPr");
  if (!tblPr) return null;

  // Check for table cell margins (w:tblCellMar)
  const tblCellMar = findChild(tblPr, "w:tblCellMar");
  if (!tblCellMar) return null;

  // Parse top margin
  const top = findChild(tblCellMar, "w:top");
  if (top?.attributes["w:w"]) {
    const twentieths = parseInt(top.attributes["w:w"] as string);
    if (!isNaN(twentieths)) {
      props.marginTop = twentieths;
    }
  }

  // Parse bottom margin
  const bottom = findChild(tblCellMar, "w:bottom");
  if (bottom?.attributes["w:w"]) {
    const twentieths = parseInt(bottom.attributes["w:w"] as string);
    if (!isNaN(twentieths)) {
      props.marginBottom = twentieths;
    }
  }

  // Parse left margin
  const left = findChild(tblCellMar, "w:left");
  if (left?.attributes["w:w"]) {
    const twentieths = parseInt(left.attributes["w:w"] as string);
    if (!isNaN(twentieths)) {
      props.marginLeft = twentieths;
    }
  }

  // Parse right margin
  const right = findChild(tblCellMar, "w:right");
  if (right?.attributes["w:w"]) {
    const twentieths = parseInt(right.attributes["w:w"] as string);
    if (!isNaN(twentieths)) {
      props.marginRight = twentieths;
    }
  }

  // Return null if no margins found
  if (
    props.marginTop === undefined &&
    props.marginBottom === undefined &&
    props.marginLeft === undefined &&
    props.marginRight === undefined
  ) {
    return null;
  }

  return props;
}

/**
 * Get row properties (rowHeight)
 */
export function parseRowProperties(rowNode: Element): {
  rowHeight: string | null;
} | null {
  const props = {
    rowHeight: null as string | null,
  };

  const trPr = findChild(rowNode, "w:trPr");
  if (!trPr) return props;

  // Check for row height
  const trHeight = findChild(trPr, "w:trHeight");
  if (trHeight?.attributes["w:val"]) {
    const twips = parseInt(trHeight.attributes["w:val"] as string);
    // Convert twips to pixels (1 inch = 1440 twips = 96 pixels at 96 DPI)
    const pixels = Math.round(twips / 15);
    props.rowHeight = `${pixels}px`;
  }

  return props;
}

/**
 * Get cell properties (colSpan, rowSpan, colWidth, backgroundColor, verticalAlign, borders)
 */
export function parseCellProperties(cellNode: Element): {
  colSpan: number;
  rowSpan: number;
  colWidth: number | null;
  backgroundColor?: string;
  verticalAlign?: string;
  borderTop?: TableCellBorder;
  borderBottom?: TableCellBorder;
  borderLeft?: TableCellBorder;
  borderRight?: TableCellBorder;
} | null {
  const props: {
    colSpan: number;
    rowSpan: number;
    colWidth: number | null;
    backgroundColor?: string;
    verticalAlign?: string;
    borderTop?: TableCellBorder;
    borderBottom?: TableCellBorder;
    borderLeft?: TableCellBorder;
    borderRight?: TableCellBorder;
  } = {
    colSpan: 1,
    rowSpan: 1,
    colWidth: null as number | null,
  };

  const tcPr = findChild(cellNode, "w:tcPr");
  if (!tcPr) {
    return props;
  }

  // Check for gridSpan (colSpan)
  const gridSpan = findChild(tcPr, "w:gridSpan");
  if (gridSpan?.attributes["w:val"]) {
    props.colSpan = parseInt(gridSpan.attributes["w:val"] as string);
  }

  // Check for vMerge (rowSpan)
  const vMerge = findChild(tcPr, "w:vMerge");
  if (vMerge?.attributes["w:val"] === "continue") {
    props.rowSpan = 0;
  }

  // Check for column width
  const tcW = findChild(tcPr, "w:tcW");
  if (tcW?.attributes["w:w"]) {
    const twips = parseInt(tcW.attributes["w:w"] as string);
    props.colWidth = Math.round(twips / 15);
  }

  // Check for background color
  const shd = findChild(tcPr, "w:shd");
  if (shd?.attributes["w:fill"]) {
    props.backgroundColor = `#${shd.attributes["w:fill"]}`;
  }

  // Check for vertical alignment
  const vAlign = findChild(tcPr, "w:vAlign");
  if (vAlign?.attributes["w:val"]) {
    props.verticalAlign = vAlign.attributes["w:val"] as string;
  }

  // Check for cell borders
  const tcBorders = findChild(tcPr, "w:tcBorders");
  if (tcBorders) {
    const topBorder = parseBorder(findChild(tcBorders, "w:top") as Element);
    if (topBorder) props.borderTop = topBorder;

    const bottomBorder = parseBorder(findChild(tcBorders, "w:bottom") as Element);
    if (bottomBorder) props.borderBottom = bottomBorder;

    const leftBorder = parseBorder(findChild(tcBorders, "w:left") as Element);
    if (leftBorder) props.borderLeft = leftBorder;

    const rightBorder = parseBorder(findChild(tcBorders, "w:right") as Element);
    if (rightBorder) props.borderRight = rightBorder;
  }

  return props;
}
