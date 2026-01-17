import type { Element, Text } from "xast";
import type { DocxImportOptions } from "../option";
import type { ImageFloatingOptions } from "@docen/tiptap-extensions/types";
import type { StyleInfo } from "../parsing/styles";
import { imageMeta } from "image-meta";
import { cropImageIfNeeded, type CropRect } from "../utils/image";
import { findChild, findDeepChild, findDeepChildren } from "../utils/xml";
import { uint8ArrayToBase64 } from "../utils/base64";

/**
 * Extract all text runs from paragraph
 */
export async function extractRuns(
  paragraph: Element,
  params: {
    hyperlinks: Map<string, string>;
    images: Map<string, string>;
    options?: DocxImportOptions;
    styleInfo?: StyleInfo;
  },
): Promise<
  Array<{
    type: string;
    text?: string;
    marks?: Array<{ type: string; attrs?: Record<string, any> }>;
  }>
> {
  const { hyperlinks, images: _images, options: _options, styleInfo } = params;
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
              const image = await extractSingleImage(drawing, params);
              if (image) {
                runs.push(image);
                continue;
              }

              // 如果单张图片提取失败，尝试分组图片提取
              const imageList = await extractImages(drawing, params);
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
            const marks = extractMarks(run, styleInfo);
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
        const imageList = await extractImages(drawing, params);
        runs.push(...imageList);
        if (imageList.length > 0) {
          continue;
        }
      }

      // Check for hard break first (before checking for text)
      const br = findChild(run, "w:br");
      if (br) {
        // Extract formatting marks for hardBreak
        const marks = extractMarks(run, styleInfo);
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
      const marks = extractMarks(run, styleInfo);

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
      const px = Math.round((mergedFormat.fontSize / 1.5) * 10) / 10;
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
async function extractSingleImage(
  drawing: Element,
  params: {
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
): Promise<{
  type: string;
  attrs: {
    src: string;
    alt: string;
    width?: number;
    height?: number;
    title?: string;
    [key: string]: any;
  };
} | null> {
  const { images, options } = params;

  // Find blip (image data reference)
  const blip = findDeepChild(drawing, "a:blip");
  if (!blip?.attributes["r:embed"]) return null;

  const rId = blip.attributes["r:embed"] as string;
  let src = images.get(rId);

  if (!src) return null;

  // Extract crop rectangle from a:srcRect (DOCX unit: 1/100000 of percentage)
  const srcRect = findDeepChild(drawing, "a:srcRect");
  let crop: CropRect | undefined;

  if (srcRect) {
    const left = srcRect.attributes["l"];
    const top = srcRect.attributes["t"];
    const right = srcRect.attributes["r"];
    const bottom = srcRect.attributes["b"];

    if (left || top || right || bottom) {
      crop = {
        left: left ? parseInt(left as string, 10) : undefined,
        top: top ? parseInt(top as string, 10) : undefined,
        right: right ? parseInt(right as string, 10) : undefined,
        bottom: bottom ? parseInt(bottom as string, 10) : undefined,
      };
    }
  }

  // Apply crop if present
  if (crop && src.startsWith("data:")) {
    // Extract base64 data from data URL
    const [metadata, base64Data] = src.split(",");
    if (base64Data) {
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }

      try {
        // Crop the image
        const croppedData = await cropImageIfNeeded(bytes, crop, {
          canvasImport: options?.canvasImport,
          enabled: options?.enableImageCrop !== false,
        });

        // Convert back to data URL using optimized base64 encoder
        const croppedBase64 = uint8ArrayToBase64(croppedData);
        src = `${metadata},${croppedBase64}`;
      } catch (error) {
        // Crop failed, use original image
        console.warn("Image cropping failed, using original image:", error);
      }
    }
  }

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

  // Extract floating positioning from wp:positionH and wp:positionV
  const positionH = findDeepChild(drawing, "wp:positionH");
  const positionV = findDeepChild(drawing, "wp:positionV");
  let floating: ImageFloatingOptions | undefined;

  if (positionH || positionV) {
    floating = {
      horizontalPosition: {
        relative: (positionH?.attributes["relativeFrom"] as any) || "page",
      },
      verticalPosition: {
        relative: (positionV?.attributes["relativeFrom"] as any) || "page",
      },
    };

    if (positionH && floating) {
      const relative = positionH.attributes["relativeFrom"] as any;
      const alignEl = findChild(positionH, "wp:align");
      const offsetEl = findChild(positionH, "wp:posOffset");

      floating.horizontalPosition = {
        relative,
        ...(alignEl &&
          alignEl.children[0]?.type === "text" && {
            align: alignEl.children[0].value as any,
          }),
        ...(offsetEl &&
          offsetEl.children[0]?.type === "text" && {
            offset: parseInt(offsetEl.children[0].value, 10),
          }),
      };
    }

    if (positionV && floating) {
      const relative = positionV.attributes["relativeFrom"] as any;
      const alignEl = findChild(positionV, "wp:align");
      const offsetEl = findChild(positionV, "wp:posOffset");

      floating.verticalPosition = {
        relative,
        ...(alignEl &&
          alignEl.children[0]?.type === "text" && {
            align: alignEl.children[0].value as any,
          }),
        ...(offsetEl &&
          offsetEl.children[0]?.type === "text" && {
            offset: parseInt(offsetEl.children[0].value, 10),
          }),
      };
    }

    // Note: effectExtent and margins are not part of ImageFloatingOptions type
    // These properties exist in DOCX but are not exposed in the public API
  }

  // Extract outline from pic:spPr/a:ln
  const spPr = findDeepChild(drawing, "pic:spPr");
  let outline:
    | {
        type: "solidFill";
        solidFillType: "rgb";
        value: string;
      }
    | undefined;

  if (spPr) {
    const ln = findDeepChild(spPr, "a:ln");
    if (ln) {
      const solidFill = findDeepChild(ln, "a:solidFill");
      if (solidFill) {
        const srgbClr = findDeepChild(solidFill, "a:srgbClr");
        if (srgbClr) {
          const val = srgbClr.attributes["val"] as string;
          outline = {
            type: "solidFill",
            solidFillType: "rgb",
            value: val,
          };
        }
      }
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
      ...(floating && { floating }),
      ...(outline && { outline }),
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
async function extractImages(
  drawing: Element,
  params: {
    images: Map<string, string>;
    options?: DocxImportOptions;
  },
): Promise<
  Array<{
    type: string;
    attrs: {
      src: string;
      alt: string;
      width?: number;
      height?: number;
      title?: string;
      [key: string]: any;
    };
  }>
> {
  const { images, options: _options } = params;
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
      const pictures = findDeepChildren(groupSp, "pic:pic");
      const pictures2 = findDeepChildren(groupSp, "pic");

      allPictures = [...pictures, ...pictures2];
    } else {
      // Some grouped images have pic:pic as direct children of wpg:wgp
      const directPictures = findDeepChildren(group, "pic:pic");
      const directPictures2 = findDeepChildren(group, "pic");

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

      const image = await extractSingleImage(syntheticDrawing, params);
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
    const image = await extractSingleImage(drawing, params);
    if (image) {
      result.push(image);
    }
  }

  return result;
}
