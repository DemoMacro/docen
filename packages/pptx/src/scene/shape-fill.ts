/** Office fill options → the renderer-native paint carried by drawing
 *  members: solid hex, Leafer gradients, or repeatable/stretched images. */

import { outlineOf } from "@docen/core/geometry";
import type { LayoutDrawingFill, LayoutDrawingLine } from "@docen/layout";
import type {
  ColorTransformOptions,
  FillOptions,
  GradientFillOptions,
  OutlineFillProperties,
  OutlineOptions,
  SolidFillOptions,
} from "@office-open/core/drawing";
import type { ColorMappingOptions } from "@office-open/core/theme";

import { tilePaintOf } from "./blip-tile";
import { transformColor } from "./color-transform";
import { patternTileSrcOf } from "./pattern-background";
import { pictureSrcOf } from "./pictures";
import type { ThemeColors } from "./table-style";

interface FillContext {
  themeColors?: ThemeColors;
  colorMapping?: ColorMappingOptions;
}

/** A raw EG_ColorChoice (a hex string or an office-open color option) →
 *  base hex (transforms ignored). */
function baseColorOf(
  value: string | SolidFillOptions | undefined,
  context: FillContext,
): string | undefined {
  if (typeof value === "string") return value.replace("#", "").toUpperCase();
  if (!value) return undefined;
  // scRGB/hsl carry channels instead of a token — they resolve to nothing here.
  if (!("value" in value)) return undefined;
  const token = value.value;
  if (/^[0-9A-F]{6}$/i.test(token)) return token.toUpperCase();
  const mapping = context.colorMapping;
  const key = LONG_COLOR_TOKEN[token] ?? token;
  const slot = mapping?.[key as keyof ColorMappingOptions] ?? token;
  return context.themeColors?.[slot as keyof ThemeColors];
}

/** The clrMap's short slot spellings → the parsed color map's long keys. */
const LONG_COLOR_TOKEN: Record<string, string> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
};

function transformsOf(
  value: string | SolidFillOptions | undefined,
): ColorTransformOptions | undefined {
  return typeof value === "object" ? value.transforms : undefined;
}

function colorOf(
  value: string | SolidFillOptions | undefined,
  context: FillContext,
): string | undefined {
  const base = baseColorOf(value, context);
  return base ? transformColor(base, transformsOf(value)).color : undefined;
}

function alphaOf(
  value: string | SolidFillOptions | undefined,
  context: FillContext = {},
): number | undefined {
  const base = baseColorOf(value, context);
  if (!base) return undefined;
  return transformColor(base, transformsOf(value)).alpha;
}

function colorPaintOf(
  value: string | SolidFillOptions | undefined,
  context: FillContext,
): string | undefined {
  const hex = colorOf(value, context);
  if (!hex) return undefined;
  const alpha = alphaOf(value, context);
  if (alpha == null || alpha >= 1) return `#${hex}`;
  const raw = Number.parseInt(hex, 16);
  return `rgba(${(raw >> 16) & 255}, ${(raw >> 8) & 255}, ${raw & 255}, ${alpha})`;
}

function gradientPaintOf(
  fill: Extract<FillOptions, { type: "gradient" }>,
  width: number,
  height: number,
  context: FillContext,
): LayoutDrawingFill | undefined {
  const shorthand = "options" in fill ? undefined : fill;
  const options: GradientFillOptions = "options" in fill ? fill.options : fill;
  const stops = options.stops
    .map((stop) => ({
      offset: stop.position / 100,
      color: colorPaintOf(stop.color, context),
    }))
    .filter((stop): stop is { offset: number; color: string } => stop.color != null);
  if (stops.length < 2) return undefined;
  const shade = options.shade;
  const shadeAngle = shade && "angle" in shade ? shade.angle : undefined;
  const path = shorthand?.path ?? (shade && "path" in shade ? shade.path : undefined);
  const angle = shorthand?.angle ?? shadeAngle;
  if (path) {
    return {
      type: "radial",
      from: { x: width / 2, y: height / 2 },
      to: { x: width / 2, y: height },
      stops,
    };
  }
  const theta = ((angle ?? 0) * Math.PI) / 180;
  const length = width * Math.abs(Math.cos(theta)) + height * Math.abs(Math.sin(theta));
  return {
    type: "linear",
    from: {
      x: width / 2 - (Math.cos(theta) * length) / 2,
      y: height / 2 - (Math.sin(theta) * length) / 2,
    },
    to: {
      x: width / 2 + (Math.cos(theta) * length) / 2,
      y: height / 2 + (Math.sin(theta) * length) / 2,
    },
    stops,
  };
}

/** Map a shape's spPr fill; solid fills retain the old hex contract, while
 *  schema gradient/blip/pattern fills become native renderer paints. */
export function shapeFillOf(
  fill: FillOptions | null | undefined,
  width: number,
  height: number,
  context: FillContext = {},
): { fill?: LayoutDrawingFill; opacity?: number } {
  if (typeof fill === "string") return { fill: fill.replace("#", "").toUpperCase() };
  if (!fill || fill.type === "none" || fill.type === "group") return {};
  if (fill.type === "solid") {
    const color = colorOf(fill.color, context);
    if (!color) return {};
    const alpha = alphaOf(fill.color, context);
    return {
      fill: color,
      ...(alpha != null && alpha < 1 ? { opacity: alpha } : {}),
    };
  }
  if (fill.type === "gradient") return { fill: gradientPaintOf(fill, width, height, context) };
  if (fill.type === "blip") {
    const url = pictureSrcOf(fill.data, fill.imageType);
    if (!url) return {};
    if (!fill.tile) return { fill: { type: "image", url, mode: "stretch" } };
    return {
      fill: {
        type: "image",
        url,
        ...tilePaintOf(fill.tile),
      },
    };
  }
  const pattern = fill;
  const paint = patternTileSrcOf({
    ...pattern,
    ...(pattern.foregroundColor != null
      ? { foregroundColor: colorPaintOf(pattern.foregroundColor, context) }
      : {}),
    ...(pattern.backgroundColor != null
      ? { backgroundColor: colorPaintOf(pattern.backgroundColor, context) }
      : {}),
  });
  return { fill: { type: "image", url: paint, mode: "repeat", repeat: true } };
}

/** Map an a:ln fill to the same renderer-native paint used by shape fills. */
function outlinePaintOf(
  outline: OutlineFillProperties | undefined,
  width: number,
  height: number,
  context: FillContext = {},
): LayoutDrawingFill | undefined {
  if (!outline) return undefined;
  if (outline.type === "solidFill" || outline.color != null)
    return colorPaintOf(outline.color ?? "000000", context);
  if (outline.type === "gradFill" && outline.gradientFill)
    return shapeFillOf({ type: "gradient", options: outline.gradientFill }, width, height, context)
      .fill;
  if (outline.type === "pattFill" && outline.patternFill)
    return shapeFillOf({ type: "pattern", ...outline.patternFill }, width, height, context).fill;
  return undefined;
}

/** One a:ln outline as the native stroke paint the painter can hand to
 *  Leafer directly; solid strokes keep the shared line contract. */
export function outlineLineOf(
  outline: OutlineOptions | undefined,
  width: number,
  height: number,
  context: FillContext = {},
): LayoutDrawingLine | undefined {
  const line = outlineOf(outline);
  if (!line) return undefined;
  const stroke = outlinePaintOf(outline, width, height, context);
  return stroke && (typeof stroke !== "string" || stroke.startsWith("rgba"))
    ? { ...line, stroke }
    : line;
}
