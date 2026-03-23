import type { Element } from "xast";
import type { Border } from "@docen/extensions/types";
import { findChild } from "@docen/utils";
import { convertTwipToPixels } from "@docen/utils";
import { parseBorder } from "./styles";

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
    const pixels = convertTwipToPixels(twips);
    props.rowHeight = `${pixels}px`;
  }

  return props;
}

/**
 * Get cell properties (colspan, rowspan, colwidth, backgroundColor, verticalAlign, borders)
 */
export function parseCellProperties(cellNode: Element): {
  colspan: number;
  rowspan: number;
  colwidth: number[] | null;
  backgroundColor?: string;
  verticalAlign?: string;
  borderTop?: Border;
  borderBottom?: Border;
  borderLeft?: Border;
  borderRight?: Border;
} | null {
  const props: {
    colspan: number;
    rowspan: number;
    colwidth: number[] | null;
    backgroundColor?: string;
    verticalAlign?: string;
    borderTop?: Border;
    borderBottom?: Border;
    borderLeft?: Border;
    borderRight?: Border;
  } = {
    colspan: 1,
    rowspan: 1,
    colwidth: null as number[] | null,
  };

  const tcPr = findChild(cellNode, "w:tcPr");
  if (!tcPr) {
    return props;
  }

  // Check for gridSpan (colspan)
  const gridSpan = findChild(tcPr, "w:gridSpan");
  if (gridSpan?.attributes["w:val"]) {
    props.colspan = parseInt(gridSpan.attributes["w:val"] as string);
  }

  // Check for vMerge (rowspan)
  const vMerge = findChild(tcPr, "w:vMerge");
  if (vMerge) {
    const vMergeVal = vMerge.attributes["w:val"] as string | undefined;
    if (vMergeVal === "restart") {
      // This cell starts a new vertical merge
      // Explicitly set rowspan to 1 to mark it for calculateRowspan
      props.rowspan = 1;
    } else {
      // vMerge="continue" or vMerge with no value (defaults to "continue")
      // This cell is merged vertically (hidden)
      props.rowspan = 0;
    }
  }

  // Check for column width
  const tcW = findChild(tcPr, "w:tcW");
  if (tcW?.attributes["w:w"]) {
    const twips = parseInt(tcW.attributes["w:w"] as string);
    const width = convertTwipToPixels(twips);
    // DOCX stores single cell width → convert to array format
    props.colwidth = [width];
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
