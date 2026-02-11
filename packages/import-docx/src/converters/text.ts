import type { Element, Text } from "xast";
import type { ParseContext } from "../parser";
import type { StyleInfo } from "../parsers/styles";
import { findChild, TEXT_ALIGN_MAP, PIXELS_PER_HALF_POINT } from "@docen/utils";
import {
  findDrawingElement,
  extractImageFromDrawing,
  extractImagesFromDrawing,
} from "../parsers/images";

/**
 * Extract text node from run
 */
function extractTextFromRun(
  run: Element,
  styleInfo?: StyleInfo,
): {
  type: string;
  text: string;
  marks?: Array<{ type: string; attrs?: Record<string, any> }>;
} | null {
  const textElement = findChild(run, "w:t");
  if (!textElement) return null;

  const text = textElement.children.find((c): c is Text => c.type === "text");
  if (!text?.value) return null;

  const marks = extractMarks(run, styleInfo);
  return {
    type: "text",
    text: text.value,
    ...(marks.length && { marks }),
  };
}

/**
 * Extract all text runs from paragraph
 */
export async function extractRuns(
  paragraph: Element,
  params: { context: ParseContext; styleInfo?: StyleInfo },
): Promise<
  Array<{
    type: string;
    text?: string;
    marks?: Array<{ type: string; attrs?: Record<string, any> }>;
  }>
> {
  const { context, styleInfo } = params;
  const runs: Array<{
    type: string;
    text?: string;
    marks?: Array<{ type: string; attrs?: Record<string, any> }>;
  }> = [];

  for (const child of paragraph.children) {
    if (child.type !== "element") continue;

    if (child.name === "w:hyperlink") {
      const hyperlink = child as Element;
      const rId = hyperlink.attributes["r:id"] as string;
      const href = context.hyperlinks.get(rId);
      if (!href) continue;

      for (const hlChild of hyperlink.children) {
        if (hlChild.type !== "element" || hlChild.name !== "w:r") continue;

        const run = hlChild as Element;
        const drawing = findDrawingElement(run);

        if (drawing) {
          const image = await extractImageFromDrawing(drawing, { context });
          if (image) {
            runs.push(image);
            continue;
          }

          const imageList = await extractImagesFromDrawing(drawing, { context });
          if (imageList.length) {
            runs.push(...imageList);
            continue;
          }
        }

        const textNode = extractTextFromRun(run, styleInfo);
        if (textNode) {
          textNode.marks = textNode.marks || [];
          textNode.marks.push({ type: "link", attrs: { href } });
          runs.push(textNode);
        }
      }
    } else if (child.name === "w:r") {
      const run = child as Element;
      const drawing = findDrawingElement(run);

      if (drawing) {
        const imageList = await extractImagesFromDrawing(drawing, { context });
        if (imageList.length) {
          runs.push(...imageList);
          continue;
        }
      }

      const br = findChild(run, "w:br");
      if (br) {
        const marks = extractMarks(run, styleInfo);
        runs.push({ type: "hardBreak", ...(marks.length && { marks }) });
      }

      const textNode = extractTextFromRun(run, styleInfo);
      if (textNode) runs.push(textNode);
    }
  }

  return runs;
}

/**
 * Extract formatting marks
 * Merges style character format with run-level formatting (run takes precedence)
 */
export function extractMarks(
  run: Element,
  styleInfo?: StyleInfo,
): Array<{ type: string; attrs?: Record<string, any> }> {
  const marks: Array<{ type: string; attrs?: Record<string, any> }> = [];

  // Find w:rPr (run properties)
  const rPr = findChild(run, "w:rPr");

  // Step 1: Initialize with style character format (base layer)
  let mergedFormat: {
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    fontSize?: number; // Half-points
    fontFamily?: string;
    backgroundColor?: string;
  } = {};

  if (styleInfo?.charFormat) {
    mergedFormat = { ...styleInfo.charFormat };
  }

  // Step 2: Run-level format overrides (higher priority)
  if (rPr) {
    // Bold (run can override style)
    const boldEl = findChild(rPr, "w:b");
    if (boldEl) {
      // Check if explicitly set to false
      if (boldEl.attributes["w:val"] === "false") {
        mergedFormat.bold = false;
      } else {
        mergedFormat.bold = true;
      }
    }

    // Italic
    const italicEl = findChild(rPr, "w:i");
    if (italicEl) {
      if (italicEl.attributes["w:val"] === "false") {
        mergedFormat.italic = false;
      } else {
        mergedFormat.italic = true;
      }
    }

    // Underline
    if (findChild(rPr, "w:u")) {
      mergedFormat.underline = true;
    }

    // Strike
    if (findChild(rPr, "w:strike")) {
      mergedFormat.strike = true;
    }

    // Text color (run overrides style)
    const colorEl = findChild(rPr, "w:color");
    if (colorEl?.attributes["w:val"] && colorEl.attributes["w:val"] !== "auto") {
      const colorVal = colorEl.attributes["w:val"] as string;
      mergedFormat.color = colorVal.startsWith("#") ? colorVal : `#${colorVal}`;
    }

    // Font size (run overrides style)
    const szEl = findChild(rPr, "w:sz");
    if (szEl?.attributes["w:val"]) {
      const sizeVal = szEl.attributes["w:val"] as string;
      const size = parseInt(sizeVal, 10);
      if (!isNaN(size)) {
        mergedFormat.fontSize = size;
      }
    }

    // Font family (run overrides style)
    const rFontsEl = findChild(rPr, "w:rFonts");
    if (rFontsEl?.attributes["w:ascii"]) {
      mergedFormat.fontFamily = rFontsEl.attributes["w:ascii"] as string;
    }

    // Background color (shading)
    const shdEl = findChild(rPr, "w:shd");
    if (shdEl?.attributes["w:fill"] && shdEl.attributes["w:fill"] !== "auto") {
      const fillColor = shdEl.attributes["w:fill"] as string;
      mergedFormat.backgroundColor = fillColor.startsWith("#") ? fillColor : `#${fillColor}`;
    }

    // Highlight
    if (findChild(rPr, "w:highlight")) {
      marks.push({ type: "highlight" });
    }

    // Subscript/Superscript
    const vertAlign = findChild(rPr, "w:vertAlign");
    if (vertAlign) {
      const val = vertAlign.attributes["w:val"] as string;
      if (val === "subscript") {
        marks.push({ type: "subscript" });
      } else if (val === "superscript") {
        marks.push({ type: "superscript" });
      }
    }
  }

  // Step 3: Convert merged format to marks
  if (mergedFormat.bold) {
    marks.push({ type: "bold" });
  }

  if (mergedFormat.italic) {
    marks.push({ type: "italic" });
  }

  if (mergedFormat.underline) {
    marks.push({ type: "underline" });
  }

  if (mergedFormat.strike) {
    marks.push({ type: "strike" });
  }

  // Text style (colors, font size, font family, etc.)
  if (
    mergedFormat.color ||
    mergedFormat.backgroundColor ||
    mergedFormat.fontSize ||
    mergedFormat.fontFamily
  ) {
    const textStyleAttrs: Record<string, string> = {
      color: mergedFormat.color || "",
      backgroundColor: mergedFormat.backgroundColor || "",
      fontSize: "",
      fontFamily: "",
      lineHeight: "",
    };

    // Font size (convert half-points to px)
    if (mergedFormat.fontSize) {
      const px = Math.round(mergedFormat.fontSize * PIXELS_PER_HALF_POINT * 10) / 10;
      textStyleAttrs.fontSize = `${px}px`;
    }

    if (mergedFormat.fontFamily) {
      textStyleAttrs.fontFamily = mergedFormat.fontFamily;
    }

    marks.push({ type: "textStyle", attrs: textStyleAttrs });
  }

  return marks;
}

/**
 * Extract text alignment
 */
export function extractAlignment(
  paragraph: Element,
): { textAlign: "left" | "right" | "center" | "justify" } | undefined {
  // Find w:pPr > w:jc
  const pPr = findChild(paragraph, "w:pPr");
  if (!pPr) return undefined;

  const jc = findChild(pPr, "w:jc");
  if (!jc?.attributes["w:val"]) return undefined;

  const alignment = jc.attributes["w:val"] as string;
  const textAlign =
    TEXT_ALIGN_MAP.docxToTipTap[alignment as keyof typeof TEXT_ALIGN_MAP.docxToTipTap];
  return textAlign ? { textAlign } : undefined;
}
