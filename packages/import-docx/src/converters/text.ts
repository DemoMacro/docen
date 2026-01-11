import type { Element, Text } from "xast";

/**
 * Extract all text runs from paragraph
 */
export function extractRuns(
  hyperlinks: Map<string, string>,
  paragraph: Element,
  images: Map<string, string>,
): Array<{
  type: string;
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, any> }>;
}> {
  const runs: Array<{
    type: string;
    text?: string;
    marks?: Array<{ type: string; attrs?: Record<string, any> }>;
  }> = [];

  // Find all w:r (text runs) and w:hyperlink (hyperlinks) by traversing children
  for (const child of paragraph.children) {
    if (child.type !== "element") continue;

    // Handle hyperlinks
    if (child.name === "w:hyperlink") {
      const hyperlink = child as Element;
      const rId = hyperlink.attributes["r:id"] as string;
      const href = hyperlinks.get(rId);

      if (href) {
        // Process all w:r elements inside the hyperlink
        for (const hlChild of hyperlink.children) {
          if (hlChild.type === "element" && hlChild.name === "w:r") {
            const run = hlChild as Element;

            // Check for image inside hyperlink
            const drawing = findChild(run, "w:drawing");
            if (drawing) {
              const image = extractImage(drawing, images);
              if (image) {
                runs.push(image);
              }
              continue;
            }

            // Extract text
            const textElement = findChild(run, "w:t");
            if (!textElement) continue;

            const text = textElement.children.find(
              (c): c is Text => c.type === "text",
            );
            if (!text || !text.value) continue;

            // Extract formatting marks
            const marks = extractMarks(run);
            // Add link mark
            marks.push({ type: "link", attrs: { href } });

            const textNode: {
              type: string;
              text: string;
              marks?: Array<{ type: string; attrs?: Record<string, any> }>;
            } = {
              type: "text",
              text: text.value,
            };

            if (marks.length > 0) {
              textNode.marks = marks;
            }

            runs.push(textNode);
          }
        }
      }
      continue;
    }

    // Handle regular text runs
    if (child.name === "w:r") {
      const run = child as Element;

      // Check for image
      const drawing = findChild(run, "w:drawing");
      if (drawing) {
        const image = extractImage(drawing, images);
        if (image) {
          runs.push(image);
        }
        continue;
      }

      // Check for hard break first (before checking for text)
      const br = findChild(run, "w:br");
      if (br) {
        // Extract formatting marks for hardBreak
        const marks = extractMarks(run);
        const hardBreakNode: {
          type: string;
          marks?: Array<{ type: string; attrs?: Record<string, any> }>;
        } = {
          type: "hardBreak",
        };

        if (marks.length > 0) {
          hardBreakNode.marks = marks;
        }

        runs.push(hardBreakNode);
      }

      // Extract text
      // Extract text
      const textElement = findChild(run, "w:t");
      if (!textElement) continue;

      const text = textElement.children.find(
        (c): c is Text => c.type === "text",
      );
      if (!text || !text.value) continue;

      // Extract formatting marks
      const marks = extractMarks(run);

      const textNode: {
        type: string;
        text: string;
        marks?: Array<{ type: string; attrs?: Record<string, any> }>;
      } = {
        type: "text",
        text: text.value,
      };

      if (marks.length > 0) {
        textNode.marks = marks;
      }

      runs.push(textNode);
    }
  }

  return runs;
}

/**
 * Extract formatting marks
 */
export function extractMarks(
  run: Element,
): Array<{ type: string; attrs?: Record<string, any> }> {
  const marks: Array<{ type: string; attrs?: Record<string, any> }> = [];

  // Find w:rPr (run properties)
  const rPr = findChild(run, "w:rPr");
  if (!rPr) return marks;

  // Bold
  if (findChild(rPr, "w:b")) {
    marks.push({ type: "bold" });
  }

  // Italic
  if (findChild(rPr, "w:i")) {
    marks.push({ type: "italic" });
  }

  // Underline
  if (findChild(rPr, "w:u")) {
    marks.push({ type: "underline" });
  }

  // Strike
  if (findChild(rPr, "w:strike")) {
    marks.push({ type: "strike" });
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

  // Text style (colors, font size, font family, etc.)
  // Check if DOCX has any text style properties
  const hasColor = findChild(rPr, "w:color");
  const hasBackgroundColor = findChild(rPr, "w:shd");
  const hasFontSize = findChild(rPr, "w:sz");
  const hasFontFamily = findChild(rPr, "w:rFonts");

  // Only create textStyle if there's at least one style property
  // This matches TipTap HTML parser behavior
  if (hasColor || hasBackgroundColor || hasFontSize || hasFontFamily) {
    const textStyleAttrs: Record<string, string> = {
      color: "",
      backgroundColor: "",
      fontSize: "",
      fontFamily: "",
      lineHeight: "",
    };

    // Text color
    if (hasColor && hasColor.attributes["w:val"]) {
      const colorVal = hasColor.attributes["w:val"] as string;
      if (colorVal !== "auto") {
        // Convert hex color (without #) to with #
        const hexColor = colorVal.startsWith("#") ? colorVal : `#${colorVal}`;
        textStyleAttrs.color = hexColor;
      }
    }

    // Background color (shading)
    if (hasBackgroundColor && hasBackgroundColor.attributes["w:fill"]) {
      const fillColor = hasBackgroundColor.attributes["w:fill"] as string;
      if (fillColor !== "auto") {
        const hexColor = fillColor.startsWith("#")
          ? fillColor
          : `#${fillColor}`;
        textStyleAttrs.backgroundColor = hexColor;
      }
    }

    // Font size (convert half-points to px)
    if (hasFontSize && hasFontSize.attributes["w:val"]) {
      const halfPoints = hasFontSize.attributes["w:val"] as string;
      const sizeValue = parseFloat(halfPoints);
      if (!isNaN(sizeValue)) {
        // Convert half-points to px: 1 half-point = 0.5pt, 1pt ≈ 1.33px
        // So: half-points / 2 * 4/3 = half-points / 1.5
        const px = Math.round((sizeValue / 1.5) * 10) / 10; // Round to 1 decimal
        textStyleAttrs.fontSize = `${px}px`;
      }
    }

    // Font family
    if (hasFontFamily && hasFontFamily.attributes["w:ascii"]) {
      textStyleAttrs.fontFamily = hasFontFamily.attributes["w:ascii"] as string;
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
  const map: Record<string, "left" | "right" | "center" | "justify"> = {
    left: "left",
    right: "right",
    center: "center",
    both: "justify",
  };

  const textAlign = map[alignment];
  return textAlign ? { textAlign } : undefined;
}

/**
 * Extract image
 */
function extractImage(
  drawing: Element,
  images: Map<string, string>,
): {
  type: string;
  attrs: {
    src: string;
    alt: string;
    [key: string]: any;
  };
} | null {
  // Find blip (image data reference)
  const blip = findDeepChild(drawing, "a:blip");
  if (!blip?.attributes["r:embed"]) return null;

  const rId = blip.attributes["r:embed"] as string;
  const src = images.get(rId);

  if (!src) return null;

  return {
    type: "image",
    attrs: {
      src,
      alt: "",
      // TODO: Extract width, height and other attributes
    },
  };
}

/**
 * Helper: Find first child element with given name
 */
function findChild(element: Element, name: string): Element | undefined {
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }
  }
  return undefined;
}

/**
 * Recursively find first child element with given name at any depth
 */
function findDeepChild(element: Element, name: string): Element | undefined {
  // Check direct children
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }
  }

  // Recursively search in children
  for (const child of element.children) {
    if (child.type === "element") {
      const found = findDeepChild(child as Element, name);
      if (found) return found;
    }
  }

  return undefined;
}
