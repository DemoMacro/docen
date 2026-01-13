import type { Element, Text } from "xast";
import { imageMeta } from "image-meta";

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

            // Check for image inside hyperlink (both direct and wrapped in mc:AlternateContent)
            let drawing = findChild(run, "w:drawing");

            // If not found directly, check for mc:AlternateContent > mc:Choice > w:drawing
            if (!drawing) {
              const altContent = findChild(run, "mc:AlternateContent");
              if (altContent) {
                const choice = findChild(altContent, "mc:Choice");
                if (choice) {
                  drawing = findChild(choice, "w:drawing");
                }
              }
            }

            if (drawing) {
              // 对于超链接中的图片，使用单张图片的提取逻辑（从 wp:extent 获取正确的尺寸）
              const image = extractSingleImage(drawing, images);
              if (image) {
                runs.push(image);
                continue;
              }

              // 如果单张图片提取失败，尝试分组图片提取
              const imageList = extractImages(drawing, images);
              runs.push(...imageList);
              if (imageList.length > 0) {
                continue;
              }
            }

            // Extract text
            const textElement = findChild(run, "w:t");
            if (!textElement) continue;

            const text = textElement.children.find((c): c is Text => c.type === "text");
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

      // Check for image (both direct and wrapped in mc:AlternateContent)
      let drawing = findChild(run, "w:drawing");

      // If not found directly, check for mc:AlternateContent > mc:Choice > w:drawing
      if (!drawing) {
        const altContent = findChild(run, "mc:AlternateContent");
        if (altContent) {
          const choice = findChild(altContent, "mc:Choice");
          if (choice) {
            drawing = findChild(choice, "w:drawing");
          }
        }
      }

      if (drawing) {
        const imageList = extractImages(drawing, images);
        runs.push(...imageList);
        if (imageList.length > 0) {
          continue;
        }
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

      const text = textElement.children.find((c): c is Text => c.type === "text");
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
export function extractMarks(run: Element): Array<{ type: string; attrs?: Record<string, any> }> {
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
        const hexColor = fillColor.startsWith("#") ? fillColor : `#${fillColor}`;
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
 * Extract single image from a drawing element
 * @param drawing - The drawing element containing the image
 * @param images - Map of relationship IDs to image data URLs
 * @returns Single image node or null
 */
function extractSingleImage(
  drawing: Element,
  images: Map<string, string>,
): {
  type: string;
  attrs: {
    src: string;
    alt: string;
    width?: number;
    height?: number;
    title?: string;
    [key: string]: any;
  };
} | null {
  // Find blip (image data reference)
  const blip = findDeepChild(drawing, "a:blip");
  if (!blip?.attributes["r:embed"]) return null;

  const rId = blip.attributes["r:embed"] as string;
  const src = images.get(rId);

  if (!src) return null;

  // Extract width and height from wp:extent (EMU units: 1 inch = 914400 EMU)
  // At 96 DPI: 1 pixel = 9525 EMU
  const extent = findDeepChild(drawing, "wp:extent");
  let width: number | undefined;
  let height: number | undefined;

  if (extent) {
    const cx = extent.attributes["cx"];
    const cy = extent.attributes["cy"];

    if (typeof cx === "string") {
      const emuWidth = parseInt(cx, 10);
      if (!isNaN(emuWidth)) {
        width = Math.round(emuWidth / 9525);
      }
    }

    if (typeof cy === "string") {
      const emuHeight = parseInt(cy, 10);
      if (!isNaN(emuHeight)) {
        height = Math.round(emuHeight / 9525);
      }
    }
  }

  // Extract rotation from a:xfrm/@rot (unit: 1/60000 degrees)
  const xfrm = findDeepChild(drawing, "a:xfrm");
  let rotation: number | undefined;

  if (xfrm) {
    const rotAttr = xfrm.attributes["rot"];
    if (typeof rotAttr === "string") {
      const rot = parseInt(rotAttr, 10);
      if (!isNaN(rot)) {
        // Convert from 1/60000 degrees to degrees
        rotation = rot / 60000;
      }
    }
  }

  // Extract title from wp:docPr
  const docPr = findDeepChild(drawing, "wp:docPr");
  let title: string | undefined;

  if (docPr) {
    const titleAttr = docPr.attributes["title"];
    if (typeof titleAttr === "string" && titleAttr) {
      title = titleAttr;
    }
  }

  return {
    type: "image",
    attrs: {
      src,
      alt: "",
      ...(width !== undefined && { width }),
      ...(height !== undefined && { height }),
      ...(rotation !== undefined && { rotation }),
      ...(title !== undefined && { title }),
    },
  };
}

/**
 * Extract images from a drawing element
 * Handles both single images and grouped images (<wpg:wgp>)
 * @param drawing - The drawing element
 * @param images - Map of relationship IDs to image data URLs
 * @returns Array of image nodes (may be empty if no images found)
 */
function extractImages(
  drawing: Element,
  images: Map<string, string>,
): Array<{
  type: string;
  attrs: {
    src: string;
    alt: string;
    width?: number;
    height?: number;
    title?: string;
    [key: string]: any;
  };
}> {
  const result: Array<{
    type: string;
    attrs: {
      src: string;
      alt: string;
      width?: number;
      height?: number;
      title?: string;
      [key: string]: any;
    };
  }> = [];

  // Check for grouped images (<wpg:wgp>)
  // The wpg:wgp can be a direct child of a:graphicData
  const inline = findChild(drawing, "wp:inline") || findChild(drawing, "wp:anchor");
  if (!inline) {
    return result;
  }

  // Get group-level dimensions from wp:extent for grouped images
  // This represents the display size of the entire group
  const extent = findChild(inline, "wp:extent");
  let groupWidth: number | undefined;
  let groupHeight: number | undefined;

  if (extent) {
    const cx = extent.attributes["cx"];
    const cy = extent.attributes["cy"];

    if (typeof cx === "string") {
      const emuWidth = parseInt(cx, 10);
      if (!isNaN(emuWidth)) {
        groupWidth = Math.round(emuWidth / 9525);
      }
    }

    if (typeof cy === "string") {
      const emuHeight = parseInt(cy, 10);
      if (!isNaN(emuHeight)) {
        groupHeight = Math.round(emuHeight / 9525);
      }
    }
  }

  const graphic = findChild(inline, "a:graphic");
  if (!graphic) {
    return result;
  }

  const graphicData = findChild(graphic, "a:graphicData");
  if (!graphicData) {
    return result;
  }

  // Check if graphicData contains wpg:wgp (grouped image)
  const group = findChild(graphicData, "wpg:wgp");

  if (group) {
    // Find all <pic:pic> elements within the group (with or without namespace)
    const groupSp = findChild(group, "wpg:grpSp");

    let allPictures: Element[] = [];

    if (groupSp) {
      // Collect all picture elements from the group
      // Try both "pic:pic" and just "pic" (namespace might be stripped)
      const pictures = findChildrenRecursive(groupSp, "pic:pic");
      const pictures2 = findChildrenRecursive(groupSp, "pic");

      allPictures = [...pictures, ...pictures2];
    } else {
      // Some grouped images have pic:pic as direct children of wpg:wgp
      const directPictures = findChildrenRecursive(group, "pic:pic");
      const directPictures2 = findChildrenRecursive(group, "pic");

      allPictures = [...directPictures, ...directPictures2];
    }

    // Extract each picture as a separate image
    for (const pic of allPictures) {
      // For grouped images, we need to find the graphic data
      const picGraphic = findChild(pic, "a:graphic");
      if (!picGraphic) {
        // The pic element might have blipFill directly
        const blipFill = findChild(pic, "pic:blipFill") || findDeepChild(pic, "a:blipFill");
        if (blipFill) {
          // Try to find the blip element
          const blip = findChild(blipFill, "a:blip") || findDeepChild(blipFill, "a:blip");
          if (blip && blip.attributes["r:embed"]) {
            const rId = blip.attributes["r:embed"] as string;
            const src = images.get(rId);
            if (src) {
              // For grouped images, adjust dimensions based on image metadata aspect ratio
              let adjustedWidth = groupWidth;
              let adjustedHeight = groupHeight;

              if (src && groupWidth && groupHeight) {
                try {
                  // Extract metadata to get original aspect ratio
                  let metaWidth: number | undefined;
                  let metaHeight: number | undefined;

                  if (src.startsWith("data:")) {
                    const base64Data = src.split(",")[1];
                    if (base64Data) {
                      const binaryString = atob(base64Data);
                      const bytes = new Uint8Array(binaryString.length);
                      for (let i = 0; i < binaryString.length; i++) {
                        bytes[i] = binaryString.charCodeAt(i);
                      }
                      const meta = imageMeta(bytes);
                      metaWidth = meta.width;
                      metaHeight = meta.height;
                    }
                  }

                  // Adjust dimensions based on aspect ratio
                  if (metaWidth && metaHeight) {
                    const metaRatio = metaWidth / metaHeight;
                    const groupRatio = groupWidth / groupHeight;

                    // If aspect ratios differ significantly, adjust to fit within group bounds
                    if (Math.abs(metaRatio - groupRatio) > 0.1) {
                      if (metaRatio > groupRatio) {
                        // Image is wider than group: fit to width, adjust height
                        adjustedWidth = groupWidth;
                        adjustedHeight = Math.round(groupWidth / metaRatio);
                      } else {
                        // Image is taller than group: fit to height, adjust width
                        adjustedHeight = groupHeight;
                        adjustedWidth = Math.round(groupHeight * metaRatio);
                      }
                    }
                  }
                } catch (error) {
                  // If metadata extraction fails, use group dimensions as-is
                  console.warn(`Failed to extract image metadata for aspect ratio:`, error);
                }
              }

              result.push({
                type: "image",
                attrs: {
                  src,
                  alt: "",
                  // Use adjusted dimensions to preserve aspect ratio
                  ...(adjustedWidth !== undefined && { width: adjustedWidth }),
                  ...(adjustedHeight !== undefined && { height: adjustedHeight }),
                },
              });
            }
          }
        }
        continue;
      }

      // Create a synthetic drawing element for this picture
      const syntheticDrawing = {
        type: "element",
        name: "w:drawing",
        children: [picGraphic],
        attributes: {},
      } as Element;

      const image = extractSingleImage(syntheticDrawing, images);
      if (image) {
        // For grouped images, adjust dimensions based on aspect ratio
        if (groupWidth !== undefined && groupHeight !== undefined && image.attrs.src) {
          try {
            const src = image.attrs.src;
            let metaWidth: number | undefined;
            let metaHeight: number | undefined;

            if (src.startsWith("data:")) {
              const base64Data = src.split(",")[1];
              if (base64Data) {
                const binaryString = atob(base64Data);
                const bytes = new Uint8Array(binaryString.length);
                for (let i = 0; i < binaryString.length; i++) {
                  bytes[i] = binaryString.charCodeAt(i);
                }
                const meta = imageMeta(bytes);
                metaWidth = meta.width;
                metaHeight = meta.height;
              }
            }

            // Adjust dimensions based on aspect ratio
            if (metaWidth && metaHeight) {
              const metaRatio = metaWidth / metaHeight;
              const groupRatio = groupWidth / groupHeight;

              if (Math.abs(metaRatio - groupRatio) > 0.1) {
                if (metaRatio > groupRatio) {
                  // Image is wider: fit to width
                  image.attrs.width = groupWidth;
                  image.attrs.height = Math.round(groupWidth / metaRatio);
                } else {
                  // Image is taller: fit to height
                  image.attrs.height = groupHeight;
                  image.attrs.width = Math.round(groupHeight * metaRatio);
                }
              } else {
                // Aspect ratios match, use group dimensions
                image.attrs.width = groupWidth;
                image.attrs.height = groupHeight;
              }
            } else {
              // No metadata, use group dimensions
              image.attrs.width = groupWidth;
              image.attrs.height = groupHeight;
            }
          } catch {
            // If metadata extraction fails, use group dimensions as-is
            image.attrs.width = groupWidth;
            image.attrs.height = groupHeight;
          }
        }
        result.push(image);
      }
    }
  } else {
    // Handle single image
    const image = extractSingleImage(drawing, images);
    if (image) {
      result.push(image);
    }
  }

  return result;
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

/**
 * Recursively find all child elements with given name at any depth
 * @param element - The element to search in
 * @param name - The element name to find
 * @returns Array of matching elements
 */
function findChildrenRecursive(element: Element, name: string): Element[] {
  const result: Element[] = [];

  // Check direct children
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      result.push(child);
    }
  }

  // Recursively search in children
  for (const child of element.children) {
    if (child.type === "element") {
      result.push(...findChildrenRecursive(child as Element, name));
    }
  }

  return result;
}
