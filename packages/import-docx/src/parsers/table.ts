import type { Element } from "xast";
import type { Border } from "@docen/extensions/types";
import { findChild, TWIPS_PER_INCH, DOCX_DPI } from "@docen/utils";
import { parseBorder } from "./styles";

/**
 * Get table properties (cell margins)
 */
export function parseTableProperties(tableNode: Element): {
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  layout?: "autofit" | "fixed";
  alignment?: "left" | "center" | "right";
  cellSpacing?: number;
} | null {
  const props = {
    marginTop: undefined as number | undefined,
    marginBottom: undefined as number | undefined,
    marginLeft: undefined as number | undefined,
    marginRight: undefined as number | undefined,
    layout: undefined as "autofit" | "fixed" | undefined,
    alignment: undefined as "left" | "center" | "right" | undefined,
    cellSpacing: undefined as number | undefined,
  };

  const tblPr = findChild(tableNode, "w:tblPr");
  if (!tblPr) return null;

  // Parse table layout (w:tblLayout)
  const tblLayout = findChild(tblPr, "w:tblLayout");
  if (tblLayout?.attributes["w:type"]) {
    const layoutType = tblLayout.attributes["w:type"] as string;
    if (layoutType === "autofit" || layoutType === "fixed") {
      props.layout = layoutType;
    }
  }

  // Parse table alignment (w:jc)
  const jc = findChild(tblPr, "w:jc");
  if (jc?.attributes["w:val"]) {
    const jcVal = jc.attributes["w:val"] as string;
    if (jcVal === "left" || jcVal === "center" || jcVal === "right") {
      props.alignment = jcVal;
    }
  }

  // Parse cell spacing (w:tblCellSpacing)
  const cellSpacing = findChild(tblPr, "w:tblCellSpacing");
  if (cellSpacing?.attributes["w:w"]) {
    const spacingVal = parseInt(cellSpacing.attributes["w:w"] as string);
    if (!isNaN(spacingVal)) {
      props.cellSpacing = spacingVal;
    }
  }

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
  const left = findChild(tblCellMar, "w:left") || findChild(tblCellMar, "w:start");
  if (left?.attributes["w:w"]) {
    const twentieths = parseInt(left.attributes["w:w"] as string);
    if (!isNaN(twentieths)) {
      props.marginLeft = twentieths;
    }
  }

  // Parse right margin
  const right = findChild(tblCellMar, "w:right") || findChild(tblCellMar, "w:end");
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
    props.marginRight === undefined &&
    props.layout === undefined &&
    props.alignment === undefined &&
    props.cellSpacing === undefined
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
  header?: boolean;
} | null {
  const props = {
    rowHeight: null as string | null,
    rowHeightRule: null as string | null,
    header: undefined as boolean | undefined,
  };

  const trPr = findChild(rowNode, "w:trPr");
  if (!trPr) return props;

  const trHeight = findChild(trPr, "w:trHeight");
  if (trHeight?.attributes["w:val"]) {
    const twips = parseInt(trHeight.attributes["w:val"] as string);
    // Preserve full precision (no Math.round) to avoid round-trip drift
    const pixels = (twips * DOCX_DPI) / TWIPS_PER_INCH;
    props.rowHeight = `${Math.round(pixels * 10) / 10}px`;
    // Preserve hRule ("exact" or "atLeast"); OOXML default is "atLeast"
    const hRule = trHeight.attributes["w:hRule"] as string | undefined;
    if (hRule) {
      props.rowHeightRule = hRule;
    }
  }

  const tblHeader = findChild(trPr, "w:tblHeader");
  if (tblHeader) {
    props.header = true;
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
  noWrap?: boolean;
  textDirection?: string;
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
    noWrap?: boolean;
    textDirection?: string;
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
    const wVal = parseInt(tcW.attributes["w:w"] as string);
    const wType = tcW.attributes["w:type"] as string;

    if (wType === "pct") {
      // Percentage: value is 1/50 of a percent (e.g., 5000 = 100%)
      props.colwidth = [Math.round(wVal / 50)];
    } else if (wType === "auto" || wType === "nil") {
      // Automatic sizing — don't set explicit width
    } else {
      // dxa (twips) or default
      if (!isNaN(wVal)) {
        // Preserve fractional precision to avoid round-trip drift
        const px = (wVal * DOCX_DPI) / TWIPS_PER_INCH;
        props.colwidth = [Math.round(px * 10) / 10];
      }
    }
  }

  // Check for background color
  const shd = findChild(tcPr, "w:shd");
  if (shd?.attributes["w:fill"] && shd.attributes["w:fill"] !== "auto") {
    props.backgroundColor = `#${shd.attributes["w:fill"]}`;
  }

  // Check for vertical alignment (map DOCX values to CSS equivalents)
  const vAlign = findChild(tcPr, "w:vAlign");
  if (vAlign?.attributes["w:val"]) {
    const vAlignMap: Record<string, string> = {
      top: "top",
      center: "middle",
      bottom: "bottom",
      both: "middle",
    };
    props.verticalAlign = vAlignMap[vAlign.attributes["w:val"] as string];
  }

  // Check for noWrap
  const noWrap = findChild(tcPr, "w:noWrap");
  if (noWrap) {
    const val = noWrap.attributes["w:val"];
    if (val !== "0" && val !== "false") {
      props.noWrap = true;
    }
  }

  // Check for text direction
  const textDir = findChild(tcPr, "w:textDirection");
  if (textDir?.attributes["w:val"]) {
    const dirVal = textDir.attributes["w:val"] as string;
    if (dirVal === "lrTb" || dirVal === "tbRl" || dirVal === "btLr") {
      props.textDirection = dirVal;
    }
  }

  // Check for cell borders
  const tcBorders = findChild(tcPr, "w:tcBorders");
  if (tcBorders) {
    const topBorder = parseBorder(findChild(tcBorders, "w:top") as Element);
    if (topBorder) props.borderTop = topBorder;

    const bottomBorder = parseBorder(findChild(tcBorders, "w:bottom") as Element);
    if (bottomBorder) props.borderBottom = bottomBorder;

    // Left/right borders (w:left or w:start, w:right or w:end)
    const leftBorder = parseBorder(
      (findChild(tcBorders, "w:left") || findChild(tcBorders, "w:start")) as Element,
    );
    if (leftBorder) props.borderLeft = leftBorder;

    const rightBorder = parseBorder(
      (findChild(tcBorders, "w:right") || findChild(tcBorders, "w:end")) as Element,
    );
    if (rightBorder) props.borderRight = rightBorder;
  }

  return props;
}
