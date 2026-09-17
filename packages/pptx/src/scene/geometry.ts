// Shared projection primitives: the group affine map, the EMU → px helpers
// and the preset geometry classification every member domain reads.

import { measureEmu } from "@docen/core/geometry";
import { emuToPx } from "@docen/layout";
import type { GeometryGuide } from "@office-open/core";
import type { PresetGeometryOptions, ShapeType } from "@office-open/core/drawing";

/** An ancestor group chain's affine map: child boxes are computed in their
 *  own EMU space, then mapped through the group's chOff/chExt scaling. */
export interface Xform {
  sx: number;
  sy: number;
  dx: number;
  dy: number;
}
export const IDENTITY: Xform = { sx: 1, sy: 1, dx: 0, dy: 0 };

/** A measure field to px: EMU passes through (÷9525), a UM string resolves.
 *  Missing extents land at 0 — the painter clips degenerate boxes. */
export function emuOf(v: unknown): number {
  return emuToPx(measureEmu(v) ?? 0);
}

/** The geometry field's two shapes (bare token / object with adjustment
 *  guides) → the evaluator's (preset, adjustments) pair. */
export function geometryOf(g: ShapeType | PresetGeometryOptions | undefined): {
  preset?: string;
  adjustments?: readonly GeometryGuide[];
} {
  if (g == null) return {};
  if (typeof g === "string") return { preset: g };
  return { preset: g.preset, ...(g.adjustmentValues ? { adjustments: g.adjustmentValues } : {}) };
}

export const BOX_PRESETS = new Set(["rect", "roundRect", "ellipse"]);
export const STRAIGHT_PRESETS = new Set(["line", "straightConnector1"]);
