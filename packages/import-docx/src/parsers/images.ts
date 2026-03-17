import { fromXml } from "xast-util-from-xml";
import { imageMeta } from "image-meta";
import type { Element } from "xast";
import type { ImageFloatingOptions, ImageNode, ImageOutlineOptions } from "@docen/extensions/types";
import type { ImageInfo } from "../types";
import type { ParseContext } from "../parser";
import type { CropRect } from "../utils/image";
import {
  findChild,
  findDeepChild,
  findDeepChildren,
  createStringValidator,
  convertEmuStringToPixels,
} from "@docen/utils";
import { uint8ArrayToBase64, base64ToUint8Array } from "../utils/base64";
import { cropImageIfNeeded } from "../utils/image";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Type guards for valid horizontal/vertical alignment values
 */
const isValidHorizontalAlign = createStringValidator([
  "left",
  "right",
  "center",
  "inside",
  "outside",
] as const);

const isValidVerticalAlign = createStringValidator([
  "top",
  "bottom",
  "center",
  "inside",
  "outside",
] as const);

const isValidHorizontalRelative = createStringValidator([
  "page",
  "character",
  "column",
  "margin",
  "leftMargin",
  "rightMargin",
  "insideMargin",
  "outsideMargin",
] as const);

const isValidVerticalRelative = createStringValidator([
  "page",
  "paragraph",
  "margin",
  "topMargin",
  "bottomMargin",
  "insideMargin",
  "outsideMargin",
  "line",
] as const);

/**
 * Extract crop rectangle from a:srcRect element
 */
function extractCropRect(srcRect: Element): CropRect | undefined {
  const left = srcRect.attributes["l"];
  const top = srcRect.attributes["t"];
  const right = srcRect.attributes["r"];
  const bottom = srcRect.attributes["b"];

  if (!left && !top && !right && !bottom) return undefined;

  return {
    left: left ? parseInt(left as string, 10) : undefined,
    top: top ? parseInt(top as string, 10) : undefined,
    right: right ? parseInt(right as string, 10) : undefined,
    bottom: bottom ? parseInt(bottom as string, 10) : undefined,
  };
}

/**
 * Apply crop to image data and update dimensions
 * Shared logic for both direct (no picGraphic) and synthetic drawing paths
 */
async function applyCropToImage(
  pic: Element,
  imgInfo: { src: string; width?: number; height?: number },
  params: { context: ParseContext },
): Promise<{ src: string; width?: number; height?: number }> {
  // Check for crop information in pic:spPr
  const spPr = findChild(pic, "pic:spPr");
  if (!spPr || !imgInfo.src.startsWith("data:")) {
    return imgInfo;
  }

  // Use findDeepChild since srcRect might be nested (e.g., inside a:xfrm)
  const srcRect = findDeepChild(pic, "a:srcRect");
  if (!srcRect) {
    return imgInfo;
  }

  const crop = extractCropRect(srcRect);
  if (!crop || (!crop.left && !crop.top && !crop.right && !crop.bottom)) {
    return imgInfo;
  }

  try {
    const [metadata, base64Data] = imgInfo.src.split(",");
    if (!base64Data) {
      return imgInfo;
    }

    const bytes = base64ToUint8Array(base64Data);
    const croppedData = await cropImageIfNeeded(bytes, crop, {
      canvasImport: params.context.image?.canvasImport,
      enabled: params.context.image?.enableImageCrop ?? false,
    });
    const croppedBase64 = uint8ArrayToBase64(croppedData);

    // Calculate cropped dimensions
    const originalWidth = imgInfo.width || 0;
    const originalHeight = imgInfo.height || 0;
    const cropLeftPct = (crop.left || 0) / 100000;
    const cropTopPct = (crop.top || 0) / 100000;
    const cropRightPct = (crop.right || 0) / 100000;
    const cropBottomPct = (crop.bottom || 0) / 100000;

    const visibleWidthPct = 1 - cropLeftPct - cropRightPct;
    const visibleHeightPct = 1 - cropTopPct - cropBottomPct;

    const croppedWidth = Math.round(originalWidth * visibleWidthPct);
    const croppedHeight = Math.round(originalHeight * visibleHeightPct);

    return {
      src: `${metadata},${croppedBase64}`,
      width: croppedWidth,
      height: croppedHeight,
    };
  } catch (error) {
    console.warn("Grouped image cropping failed, using original image:", error);
    return imgInfo;
  }
}

/**
 * Extract horizontal position (align/offset) from position element
 */
function extractHorizontalPosition(
  positionEl: Element,
): { align?: "left" | "right" | "center" | "inside" | "outside"; offset?: number } | undefined {
  const alignEl = findChild(positionEl, "wp:align");
  const offsetEl = findChild(positionEl, "wp:posOffset");

  let align: "left" | "right" | "center" | "inside" | "outside" | undefined;
  if (alignEl?.children[0]?.type === "text") {
    const value = alignEl.children[0].value;
    if (isValidHorizontalAlign(value)) {
      align = value as "left" | "right" | "center" | "inside" | "outside";
    }
  }

  const offset =
    offsetEl?.children[0]?.type === "text" ? parseInt(offsetEl.children[0].value, 10) : undefined;

  if (!align && offset === undefined) return undefined;

  return { ...(align && { align }), ...(offset !== undefined && { offset }) };
}

/**
 * Extract vertical position (align/offset) from position element
 */
function extractVerticalPosition(
  positionEl: Element,
): { align?: "top" | "bottom" | "center" | "inside" | "outside"; offset?: number } | undefined {
  const alignEl = findChild(positionEl, "wp:align");
  const offsetEl = findChild(positionEl, "wp:posOffset");

  let align: "top" | "bottom" | "center" | "inside" | "outside" | undefined;
  if (alignEl?.children[0]?.type === "text") {
    const value = alignEl.children[0].value;
    if (isValidVerticalAlign(value)) {
      align = value as "top" | "bottom" | "center" | "inside" | "outside";
    }
  }

  const offset =
    offsetEl?.children[0]?.type === "text" ? parseInt(offsetEl.children[0].value, 10) : undefined;

  if (!align && offset === undefined) return undefined;

  return { ...(align && { align }), ...(offset !== undefined && { offset }) };
}

/**
 * Find drawing element (handles both direct and mc:AlternateContent wrapping)
 */
export function findDrawingElement(run: Element): Element | null {
  let drawing = findChild(run, "w:drawing");
  if (drawing) return drawing;

  const altContent = findChild(run, "mc:AlternateContent");
  const choice = altContent && findChild(altContent, "mc:Choice");
  return choice ? findChild(choice, "w:drawing") : null;
}

/**
 * Adjust image dimensions to fit within group bounds while preserving aspect ratio
 */
function fitToGroup(
  groupWidth: number,
  groupHeight: number,
  metaWidth: number,
  metaHeight: number,
): { width: number; height: number } {
  const metaRatio = metaWidth / metaHeight;
  const groupRatio = groupWidth / groupHeight;

  // If aspect ratios differ significantly, adjust to fit within group bounds
  if (Math.abs(metaRatio - groupRatio) > 0.1) {
    if (metaRatio > groupRatio) {
      // Image is wider: fit to width
      return { width: groupWidth, height: Math.round(groupWidth / metaRatio) };
    } else {
      // Image is taller: fit to height
      return { width: Math.round(groupHeight * metaRatio), height: groupHeight };
    }
  }

  // Aspect ratios match, use group dimensions
  return { width: groupWidth, height: groupHeight };
}

/**
 * Extract images from DOCX and convert to base64 data URLs or use custom handler
 * Returns Map of relationship ID to image info (src + dimensions)
 */
export async function extractImages(
  files: Record<string, Uint8Array>,
  handler?: import("../types").DocxImageImportHandler,
): Promise<Map<string, ImageInfo>> {
  const images = new Map<string, ImageInfo>();
  const relsXml = files["word/_rels/document.xml.rels"];
  if (!relsXml) return images;

  const relsXast = fromXml(new TextDecoder().decode(relsXml));
  const relationships = findChild(relsXast, "Relationships");
  if (!relationships) return images;

  const rels = findDeepChildren(relationships, "Relationship");
  for (const rel of rels) {
    if (rel.attributes.Type === IMAGE_REL_TYPE && rel.attributes.Id && rel.attributes.Target) {
      const imagePath = "word/" + (rel.attributes.Target as string);
      const imageData = files[imagePath];
      if (!imageData) continue;

      // Extract image metadata
      let width: number | undefined;
      let height: number | undefined;
      let imageType = "png"; // default fallback

      try {
        const meta = imageMeta(imageData);
        width = meta.width;
        height = meta.height;
        if (meta.type) imageType = meta.type;
      } catch {
        // If metadata extraction fails, use defaults
      }

      // Use custom handler or default base64 encoding
      let src: string;
      if (handler) {
        const result = await handler({
          id: rel.attributes.Id as string,
          contentType: `image/${imageType}`,
          data: imageData,
        });
        src = result.src;
      } else {
        // Default behavior: convert to base64 data URL
        const base64 = uint8ArrayToBase64(imageData);
        src = `data:image/${imageType};base64,${base64}`;
      }

      images.set(rel.attributes.Id as string, {
        src,
        width,
        height,
      });
    }
  }

  return images;
}

/**
 * Extract single image from a drawing element
 * Returns TipTap image node or null
 */
export async function extractImageFromDrawing(
  drawing: Element,
  params: { context: ParseContext },
): Promise<ImageNode | null> {
  const { context } = params;

  const blip = findDeepChild(drawing, "a:blip");
  if (!blip?.attributes["r:embed"]) return null;

  const rId = blip.attributes["r:embed"] as string;
  const imgInfo = context.images.get(rId);
  if (!imgInfo) return null;

  let src = imgInfo.src;

  // Extract and apply crop rectangle from a:srcRect (DOCX unit: 1/100000 of percentage)
  const srcRect = findDeepChild(drawing, "a:srcRect");
  if (srcRect) {
    const crop = extractCropRect(srcRect);
    if (crop && src.startsWith("data:")) {
      const [metadata, base64Data] = src.split(",");
      if (base64Data) {
        const bytes = base64ToUint8Array(base64Data);

        try {
          const croppedData = await cropImageIfNeeded(bytes, crop, {
            canvasImport: context.image?.canvasImport,
            enabled: context.image?.enableImageCrop ?? false,
          });

          const croppedBase64 = uint8ArrayToBase64(croppedData);
          src = `${metadata},${croppedBase64}`;
        } catch (error) {
          console.warn("Image cropping failed, using original image:", error);
        }
      }
    }
  }

  // Extract width and height from wp:extent
  const extent = findDeepChild(drawing, "wp:extent");
  let width: number | undefined;
  let height: number | undefined;

  if (extent) {
    const cx = extent.attributes["cx"];
    const cy = extent.attributes["cy"];

    if (typeof cx === "string") width = convertEmuStringToPixels(cx);
    if (typeof cy === "string") height = convertEmuStringToPixels(cy);
  }

  // Extract rotation from a:xfrm/@rot (unit: 1/60000 degrees)
  const xfrm = findDeepChild(drawing, "a:xfrm");
  let rotation: number | undefined;

  if (xfrm?.attributes["rot"]) {
    const rot = parseInt(xfrm.attributes["rot"] as string, 10);
    if (!isNaN(rot)) rotation = rot / 60000;
  }

  // Extract title from wp:docPr
  const docPr = findDeepChild(drawing, "wp:docPr");
  const title = docPr?.attributes["title"] as string | undefined;

  // Extract floating positioning
  const positionH = findDeepChild(drawing, "wp:positionH");
  const positionV = findDeepChild(drawing, "wp:positionV");
  let floating: ImageFloatingOptions | undefined;

  if (positionH || positionV) {
    const hPos = positionH ? extractHorizontalPosition(positionH) : undefined;
    const vPos = positionV ? extractVerticalPosition(positionV) : undefined;

    // Extract and validate relative values
    const hRelative = positionH?.attributes["relativeFrom"];
    const vRelative = positionV?.attributes["relativeFrom"];

    const horizontalRelative =
      typeof hRelative === "string" && isValidHorizontalRelative(hRelative)
        ? (hRelative as
            | "page"
            | "character"
            | "column"
            | "margin"
            | "leftMargin"
            | "rightMargin"
            | "insideMargin"
            | "outsideMargin")
        : "page";
    const verticalRelative =
      typeof vRelative === "string" && isValidVerticalRelative(vRelative)
        ? (vRelative as
            | "page"
            | "paragraph"
            | "margin"
            | "topMargin"
            | "bottomMargin"
            | "insideMargin"
            | "outsideMargin"
            | "line")
        : "page";

    floating = {
      horizontalPosition: {
        relative: horizontalRelative,
        ...(hPos?.align && { align: hPos.align }),
        ...(hPos?.offset !== undefined && { offset: hPos.offset }),
      },
      verticalPosition: {
        relative: verticalRelative,
        ...(vPos?.align && { align: vPos.align }),
        ...(vPos?.offset !== undefined && { offset: vPos.offset }),
      },
    };
  }

  // Extract outline from pic:spPr/a:ln
  const spPr = findDeepChild(drawing, "pic:spPr");
  let outline: ImageOutlineOptions | undefined;

  if (spPr) {
    const ln = findDeepChild(spPr, "a:ln");
    const solidFill = ln && findDeepChild(ln, "a:solidFill");
    const srgbClr = solidFill && findDeepChild(solidFill, "a:srgbClr");

    if (srgbClr?.attributes["val"]) {
      outline = {
        type: "solidFill",
        solidFillType: "rgb",
        value: srgbClr.attributes["val"] as string,
      };
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
      ...(title && { title }),
      ...(floating && { floating }),
      ...(outline && { outline }),
    },
  };
}

/**
 * Extract images from a drawing element
 * Handles both single images and grouped images (<wpg:wgp>)
 */
export async function extractImagesFromDrawing(
  drawing: Element,
  params: { context: ParseContext },
): Promise<ImageNode[]> {
  const result: ImageNode[] = [];

  const inline = findChild(drawing, "wp:inline") || findChild(drawing, "wp:anchor");
  if (!inline) return result;

  // Get group-level dimensions from wp:extent
  const extent = findChild(inline, "wp:extent");
  let groupWidth: number | undefined;
  let groupHeight: number | undefined;

  if (extent) {
    const cx = extent.attributes["cx"];
    const cy = extent.attributes["cy"];

    if (typeof cx === "string") groupWidth = convertEmuStringToPixels(cx);
    if (typeof cy === "string") groupHeight = convertEmuStringToPixels(cy);
  }

  const graphic = findChild(inline, "a:graphic");
  if (!graphic) return result;

  const graphicData = findChild(graphic, "a:graphicData");
  if (!graphicData) return result;

  // Check if graphicData contains wpg:wgp (grouped image)
  const group = findChild(graphicData, "wpg:wgp");

  if (group) {
    // Find all <pic:pic> elements within the group
    const groupSp = findChild(group, "wpg:grpSp");
    const pictures = groupSp
      ? [...findDeepChildren(groupSp, "pic:pic"), ...findDeepChildren(groupSp, "pic")]
      : [...findDeepChildren(group, "pic:pic"), ...findDeepChildren(group, "pic")];

    // Extract each picture as a separate image (in original XML order)
    for (const pic of pictures) {
      const picGraphic = findChild(pic, "a:graphic");

      if (!picGraphic) {
        // The pic element might have blipFill directly
        const blipFill = findChild(pic, "pic:blipFill") || findDeepChild(pic, "a:blipFill");
        if (!blipFill) continue;

        const blip = findChild(blipFill, "a:blip") || findDeepChild(blipFill, "a:blip");
        if (!blip?.attributes["r:embed"]) continue;

        const rId = blip.attributes["r:embed"] as string;
        const imgInfo = params.context.images.get(rId);
        if (!imgInfo) continue;

        // Apply crop if needed
        const processedImgInfo = await applyCropToImage(pic, imgInfo, params);

        // For grouped images, use processed image dimensions (original or cropped)
        result.push({
          type: "image",
          attrs: {
            src: processedImgInfo.src,
            alt: "",
            width: processedImgInfo.width,
            height: processedImgInfo.height,
          },
        });
        continue;
      }

      // Create a synthetic drawing element for this picture
      const syntheticDrawing = {
        type: "element",
        name: "w:drawing",
        children: [picGraphic],
        attributes: {},
      } as Element;

      const image = await extractImageFromDrawing(syntheticDrawing, params);
      if (!image) continue;

      // Check for crop information in pic:spPr (for grouped images with graphic)
      const spPr = findChild(pic, "pic:spPr");
      const srcRect = spPr ? findDeepChild(pic, "a:srcRect") : undefined;
      const hasCrop = srcRect && extractCropRect(srcRect);
      const crop = hasCrop ? extractCropRect(srcRect)! : undefined;

      if (
        crop &&
        (crop.left || crop.top || crop.right || crop.bottom) &&
        image.attrs?.src?.startsWith("data:")
      ) {
        // Apply crop
        try {
          const [metadata, base64Data] = image.attrs.src.split(",");
          if (base64Data) {
            const bytes = base64ToUint8Array(base64Data);
            const croppedData = await cropImageIfNeeded(bytes, crop, {
              canvasImport: params.context.image?.canvasImport,
              enabled: params.context.image?.enableImageCrop ?? false,
            });
            const croppedBase64 = uint8ArrayToBase64(croppedData);
            image.attrs.src = `${metadata},${croppedBase64}`;

            // Calculate cropped dimensions
            const rId =
              syntheticDrawing.children[0]?.type === "element"
                ? (findDeepChild(syntheticDrawing.children[0] as Element, "a:blip")?.attributes[
                    "r:embed"
                  ] as string)
                : undefined;

            if (rId) {
              const imgInfo = params.context.images.get(rId);
              if (imgInfo?.width && imgInfo?.height) {
                const cropLeftPct = (crop.left || 0) / 100000;
                const cropTopPct = (crop.top || 0) / 100000;
                const cropRightPct = (crop.right || 0) / 100000;
                const cropBottomPct = (crop.bottom || 0) / 100000;

                const visibleWidthPct = 1 - cropLeftPct - cropRightPct;
                const visibleHeightPct = 1 - cropTopPct - cropBottomPct;

                const croppedWidth = Math.round(imgInfo.width * visibleWidthPct);
                const croppedHeight = Math.round(imgInfo.height * visibleHeightPct);

                image.attrs.width = croppedWidth;
                image.attrs.height = croppedHeight;
              }
            }
          }
        } catch (error) {
          console.warn("Grouped image cropping failed, using original image:", error);
        }
      } else {
        // No crop, adjust dimensions based on aspect ratio
        const rId =
          syntheticDrawing.children[0]?.type === "element"
            ? (findDeepChild(syntheticDrawing.children[0] as Element, "a:blip")?.attributes[
                "r:embed"
              ] as string)
            : undefined;

        if (groupWidth && groupHeight && rId) {
          const imgInfo = params.context.images.get(rId);
          if (imgInfo?.width && imgInfo?.height) {
            const adjusted = fitToGroup(groupWidth, groupHeight, imgInfo.width, imgInfo.height);
            image.attrs!.width = adjusted.width;
            image.attrs!.height = adjusted.height;
          } else {
            image.attrs!.width = groupWidth;
            image.attrs!.height = groupHeight;
          }
        }
      }

      result.push(image);
    }
  } else {
    // Handle single image
    const image = await extractImageFromDrawing(drawing, params);
    if (image) result.push(image);
  }

  return result;
}
