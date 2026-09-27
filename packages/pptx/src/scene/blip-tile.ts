/** OOXML a:tile → the renderer-native portion of Leafer's image paint. */

import { measureEmu } from "@docen/core/geometry";
import { emuToPx } from "@docen/layout";
import type { TileOptions } from "@office-open/core/drawing";

type LeaferAlign =
  | "top-left"
  | "top"
  | "top-right"
  | "left"
  | "center"
  | "right"
  | "bottom-left"
  | "bottom"
  | "bottom-right";

const LEAFER_ALIGN: Record<NonNullable<TileOptions["alignment"]>, LeaferAlign> = {
  topLeft: "top-left",
  top: "top",
  topRight: "top-right",
  left: "left",
  center: "center",
  right: "right",
  bottomLeft: "bottom-left",
  bottom: "bottom",
  bottomRight: "bottom-right",
};

/** Leafer's paint uses ratio scales, px offsets and hyphenated aligns;
 *  office-open stores ratio-percent scales, EMU offsets and camelCase aligns.
 *  a:tile's alternating mirror mode is intentionally absent — Leafer repeat
 *  cannot express it and a negative scale would mirror every tile instead. */
export function tilePaintOf(tile: TileOptions): {
  mode: "repeat";
  repeat: true;
  scale: { x: number; y: number };
  offset?: { x: number; y: number };
  align?: LeaferAlign;
} {
  const { tx, ty, sx = 100, sy = 100, alignment } = tile;
  return {
    mode: "repeat",
    repeat: true,
    scale: { x: sx / 100, y: sy / 100 },
    ...(tx != null || ty != null
      ? {
          offset: {
            x: tx != null ? emuToPx(measureEmu(tx) ?? 0) : 0,
            y: ty != null ? emuToPx(measureEmu(ty) ?? 0) : 0,
          },
        }
      : {}),
    ...(alignment ? { align: LEAFER_ALIGN[alignment] } : {}),
  };
}
