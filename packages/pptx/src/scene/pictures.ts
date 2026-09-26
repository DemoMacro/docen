// Picture (p:pic) projection: data shapes to data URLs, source-rectangle
// crops to fractions, plus the outline/shadow dressing.

import { outlineOf, outerShadowOf } from "@docen/core/geometry";
import type { LayoutDrawingMember, LayoutPictureCrop } from "@docen/layout";
import type { SourceRectangleOptions } from "@office-open/core/drawing";
import type { PictureOptions } from "@office-open/pptx";

import { emuOf, type Xform } from "./geometry";

const PIC_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
};

/** Bytes → base64 (btoa is universal: browsers and Node ≥ 16). */
function base64Of(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

/** The base contract's three data shapes → the data URL the painter loads;
 *  emf/wmf have no browser decoder, so they resolve to undefined. */
export function pictureSrcOf(data: unknown, type: string | undefined): string | undefined {
  const mime = PIC_MIME[type ?? ""];
  return typeof data === "string"
    ? data.startsWith("data:")
      ? data
      : mime
        ? `data:${mime};base64,${data}`
        : undefined
    : mime && data instanceof Uint8Array
      ? `data:${mime};base64,${base64Of(data)}`
      : undefined;
}

/** a:srcRect integer-percent insets → the crop fractions the member carries;
 *  an all-zero rect is no crop. */
function cropOf(sr: SourceRectangleOptions | undefined): LayoutPictureCrop | undefined {
  const l = sr?.left ?? 0;
  const t = sr?.top ?? 0;
  const r = sr?.right ?? 0;
  const b = sr?.bottom ?? 0;
  if (!l && !t && !r && !b) return undefined;
  return { left: l / 100, top: t / 100, right: r / 100, bottom: b / 100 };
}

export function pictureMember(
  pic: PictureOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember | undefined {
  const crop = cropOf(pic.sourceRectangle);
  const line = outlineOf(pic.outline);
  const shadow = outerShadowOf(pic.effects);
  const src = pictureSrcOf(pic.data, pic.type);
  return {
    kind: "picture",
    x: t.sx * emuOf(pic.x) + t.dx,
    y: t.sy * emuOf(pic.y) + t.dy,
    width: t.sx * emuOf(pic.width),
    height: t.sy * emuOf(pic.height),
    ...(src ? { src } : {}),
    ...(pic.flipHorizontal ? { flipH: true } : {}),
    ...(pic.flipVertical ? { flipV: true } : {}),
    ...(pic.rotation ? { rotation: pic.rotation } : {}),
    ...(crop ? { crop } : {}),
    ...(line ? { line } : {}),
    ...(shadow ? { shadow } : {}),
    ...(childPath ? { childPath } : {}),
  };
}
