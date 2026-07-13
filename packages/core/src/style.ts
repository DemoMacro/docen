/**
 * @docen/core/style — OOXML fill/outline → LeaferJS paint properties.
 *
 * LeaferJS elements accept `fill` (color string or paint) and `stroke` /
 * `strokeWidth` / `dashPattern` (see `IStrokeStyle`). This module maps the
 * OOXML {@link FillOptions} / {@link OutlineOptions} discriminated unions onto
 * those properties, reusing the same color-normalization rules as `@docen/docx`'s
 * HTML renderer so the canvas and the HTML crop/vector fallbacks agree.
 *
 * Only solid fills / solid outlines are mapped to LeaferJS native paints in
 * stage 1; gradient / blip / pattern / effect paints return `undefined` and are
 * layered separately by the calling editor component (stage 2+).
 *
 * @module
 */

import type {
  FillOptions,
  OutlineOptions,
  PresetDash,
  SolidFillOptions,
} from "@office-open/core/drawingml";
import { convertEmuToPixels } from "@office-open/core/util";

/** A small CSS-named-color table so OOXML bare hex / named colors resolve to
 *  #RRGGBB. Mirrors `@docen/docx` `normalizeColorToHex` — kept local so
 *  @docen/core has no dependency on the docx engine. */
const CSS_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  gray: "#808080",
  grey: "#808080",
  silver: "#C0C0C0",
  maroon: "#800000",
  olive: "#808000",
  purple: "#800080",
  teal: "#008080",
  navy: "#000080",
  orange: "#FFA500",
  pink: "#FFC0CB",
  brown: "#A52A2A",
  lime: "#00FF00",
  gold: "#FFD700",
  // Keep in lock-step with @docen/docx's CSS_COLORS (extensions/utils.ts) —
  // canvas and HTML crop/vector fallbacks must resolve named colors identically
  // (edit == render). TODO: extract to a shared neutral module.
  aqua: "#00FFFF",
  fuchsia: "#FF00FF",
  indigo: "#4B0082",
  violet: "#EE82EE",
  coral: "#FF7F50",
  salmon: "#FA8072",
  tomato: "#FF6347",
};

/** Normalize any OOXML color value (bare hex, #hex, named, or { val }) to #RRGGBB.
 *  Returns undefined for "auto" (no CSS/LeaferJS equivalent) and unrecognized input. */
export const normalizeColorToHex = (color: unknown): string | undefined => {
  if (!color) return undefined;
  if (typeof color === "object") {
    const { val } = color as { val?: unknown };
    return val ? normalizeColorToHex(val) : undefined;
  }
  if (typeof color !== "string") return undefined;
  if (color === "auto") return undefined;
  if (color.startsWith("#")) {
    return color.length === 4
      ? `#${color[1]}${color[1]}${color[2]}${color[2]}${color[3]}${color[3]}`.toUpperCase()
      : color.toUpperCase();
  }
  if (/^[0-9A-Fa-f]{6}$/.test(color)) return `#${color.toUpperCase()}`;
  if (/^[0-9A-Fa-f]{3}$/.test(color)) {
    return `#${color[0]}${color[0]}${color[1]}${color[1]}${color[2]}${color[2]}`.toUpperCase();
  }
  return CSS_COLORS[color.toLowerCase()] ?? undefined;
};

/** Extract a hex color from a SolidFillOptions union member (rgb/scheme/hsl/…). */
const solidColorValue = (color: SolidFillOptions | undefined): string | undefined => {
  if (!color) return undefined;
  if (typeof color === "string") return color;
  const obj = color as { value?: string; val?: string };
  return obj.value ?? obj.val;
};

/** OOXML fill → LeaferJS `fill` paint value (a hex color string, or undefined).
 *  Solid fills only at this stage; none/gradient/blip/pattern return undefined. */
export const renderFill = (fill: FillOptions | null | undefined): string | undefined => {
  if (!fill) return undefined;
  if (typeof fill === "string") return normalizeColorToHex(fill);
  if (fill.type !== "solid") return undefined;
  const color = typeof fill.color === "string" ? fill.color : solidColorValue(fill.color);
  return normalizeColorToHex(color);
};

/** LeaferJS stroke properties (`IStrokeStyle` + `stroke` color), emitted by
 *  {@link renderOutline}. These map directly onto `IUIBaseInputData` stroke
 *  fields — no custom wrapper type. */
export interface LeaferStroke {
  /** Hex color (#RRGGBB), defaulting to black when the OOXML color is "auto". */
  stroke: string;
  /** Stroke width in pixels. */
  strokeWidth: number;
  /** Dash pattern as a number array (Canvas `setLineDash` contract). Undefined
   *  means a solid line (LeaferJS default) — never set it to a string, since
   *  `ctx.setLineDash` rejects non-numeric input. */
  dashPattern?: number[];
}

/** OOXML outline → LeaferJS stroke properties. `noFill` and absent outlines
 *  return undefined. EMU widths are converted to px via the standard 9525
 *  EMU/px ratio. The dash preset is collapsed to a `number[]` (Canvas
 *  `setLineDash` contract) for the common OOXML dot/dash presets; full
 *  PresetDash → dash-array mapping is stage 2. */
export const renderOutline = (
  outline: OutlineOptions | null | undefined,
): LeaferStroke | undefined => {
  if (!outline || outline.type === "noFill") return undefined;
  const stroke = normalizeColorToHex(solidColorValue(outline.color)) ?? "#000000";
  // 9525 EMU is OOXML's default line width (= 1 CSS px); convert via the shared
  // util so this file never hardcodes the EMU/px ratio (single source of truth
  // in @office-open/core/util, same one geometry.ts emuToPx proxies).
  const emu = typeof outline.width === "number" ? outline.width : 9525;
  const strokeWidth = Math.max(0.5, convertEmuToPixels(emu));
  const dash = outline.dash as (typeof PresetDash)[keyof typeof PresetDash] | undefined;
  // Collapse OOXML dash presets to a Canvas dash array. Solid (the common case)
  // leaves dashPattern unset so LeaferJS draws a continuous line.
  let dashPattern: number[] | undefined;
  if (dash === "sysDot" || dash === "dot") {
    dashPattern = [2, 4];
  } else if (dash === "sysDash" || dash === "dash" || dash === "sysDashDot") {
    dashPattern = [8, 4];
  }
  return { stroke, strokeWidth, dashPattern };
};
