import { fromXml } from "xast-util-from-xml";
import { imageMeta } from "image-meta";
import type { Element } from "xast";
import type { ImageFloatingOptions, ImageNode, ImageOutlineOptions } from "@docen/extensions/types";
import type { ImageInfo } from "./types";
import type { ParseContext } from "../parser";
import type { CropRect } from "../utils/image";
import { findChild, findDeepChild, findDeepChildren } from "../utils/xml";
import { uint8ArrayToBase64, base64ToUint8Array } from "../utils/base64";
import { cropImageIfNeeded } from "../utils/image";
import { convertEmuStringToPixels } from "../utils/conversion";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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
 * Extract position (align/offset) from position element
 */
function extractPosition(positionEl: Element): { align?: string; offset?: number } | undefined {
  const alignEl = findChild(positionEl, "wp:align");
  const offsetEl = findChild(positionEl, "wp:posOffset");

  const align =
    alignEl?.children[0]?.type === "text" ? (alignEl.children[0].value as string) : undefined;
  const offset =
    offsetEl?.children[0]?.type === "text" ? parseInt(offsetEl.children[0].value, 10) : undefined;

  if (!align && offset === undefined) return undefined;

  return { ...(align && { align }), ...(offset !== undefined && { offset }) };
}

/**
 * Find drawing element (handles both direct and mc:AlternateContent wrapping)
 */
export function findDrawingElement(run: Element): Element | undefined {
  let drawing = findChild(run, "w:drawing");
  if (drawing) return drawing;

  const altContent = findChild(run, "mc:AlternateContent");
  const choice = altContent && findChild(altContent, "mc:Choice");
  return choice && findChild(choice, "w:drawing");
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
 * Extract images from DOCX and convert to base64 data URLs
 * Returns Map of relationship ID to image info (src + dimensions)
 */
export function extractImages(files: Record<string, Uint8Array>): Map<string, ImageInfo> {
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

      // Extract image metadata and convert to base64 data URL in one pass
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

      // Convert to base64 data URL
      const base64 = uint8ArrayToBase64(imageData);
      const src = `data:image/${imageType};base64,${base64}`;

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
            canvasImport: context.canvasImport,
            enabled: context.enableImageCrop !== false,
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
    const hPos = positionH ? extractPosition(positionH) : undefined;
    const vPos = positionV ? extractPosition(positionV) : undefined;

    floating = {
      horizontalPosition: {
        relative: (positionH?.attributes["relativeFrom"] as any) || "page",
        ...(hPos?.align && { align: hPos.align as any }),
        ...(hPos?.offset !== undefined && { offset: hPos.offset }),
      },
      verticalPosition: {
        relative: (positionV?.attributes["relativeFrom"] as any) || "page",
        ...(vPos?.align && { align: vPos.align as any }),
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
 * Create image node with adjusted dimensions for grouped images
 */
function createGroupedImage(
  imgInfo: ImageInfo,
  groupWidth?: number,
  groupHeight?: number,
): ImageNode {
  if (groupWidth && groupHeight && imgInfo.width && imgInfo.height) {
    const adjusted = fitToGroup(groupWidth, groupHeight, imgInfo.width, imgInfo.height);
    return {
      type: "image",
      attrs: {
        src: imgInfo.src,
        alt: "",
        width: adjusted.width,
        height: adjusted.height,
      },
    };
  }

  return {
    type: "image",
    attrs: {
      src: imgInfo.src,
      alt: "",
      ...(groupWidth !== undefined && { width: groupWidth }),
      ...(groupHeight !== undefined && { height: groupHeight }),
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

    // Extract each picture as a separate image
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

        result.push(createGroupedImage(imgInfo, groupWidth, groupHeight));
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

      // For grouped images, adjust dimensions based on aspect ratio
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

      result.push(image);
    }
  } else {
    // Handle single image
    const image = await extractImageFromDrawing(drawing, params);
    if (image) result.push(image);
  }

  return result;
}
