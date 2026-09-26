/** Office fill options → the renderer-native paint carried by drawing
 *  members: solid hex, Leafer gradients, or repeatable/stretched images. */

import type { LayoutDrawingFill } from "@docen/layout";
import type { ColorTransformOptions, FillOptions } from "@office-open/core/drawing";
import type { ColorMappingOptions } from "@office-open/core/theme";

import { transformColor } from "./color-transform";
import { patternTileSrcOf } from "./pattern-background";
import { pictureSrcOf } from "./pictures";
import type { ThemeColors } from "./table-style";

interface FillContext {
  themeColors?: ThemeColors;
  colorMapping?: ColorMappingOptions;
}

/** A raw EG_ColorChoice / SolidFillOptions → base hex (transforms ignored). */
function baseColorOf(value: unknown, context: FillContext): string | undefined {
  if (typeof value === "string") return value.replace("#", "").toUpperCase();
  if (!value || typeof value !== "object") return undefined;
  const solid = value as { type?: unknown; color?: unknown; value?: unknown };
  if (solid.type === "solid") return baseColorOf(solid.color, context);
  if (typeof solid.value !== "string") return undefined;
  const token = solid.value;
  if (/^[0-9A-F]{6}$/i.test(token)) return token.toUpperCase();
  const mapping = (context.colorMapping ?? {}) as unknown as Record<string, string>;
  const key =
    { bg1: "background1", tx1: "text1", bg2: "background2", tx2: "text2" }[token] ?? token;
  const slot = mapping[key] ?? token;
  return context.themeColors?.[slot as keyof ThemeColors];
}

function transformsOf(value: unknown): ColorTransformOptions | undefined {
  if (!value || typeof value !== "object" || !("transforms" in value)) return undefined;
  const transforms = (value as { transforms?: unknown }).transforms;
  return transforms && typeof transforms === "object"
    ? (transforms as ColorTransformOptions)
    : undefined;
}

function colorOf(value: unknown, context: FillContext): string | undefined {
  const base = baseColorOf(value, context);
  return base ? transformColor(base, transformsOf(value)).color : undefined;
}

function alphaOf(value: unknown): number | undefined {
  const base = baseColorOf(value, {});
  if (!base) return undefined;
  return transformColor(base, transformsOf(value)).alpha;
}

function colorPaintOf(value: unknown, context: FillContext): string | undefined {
  const hex = colorOf(value, context);
  if (!hex) return undefined;
  const alpha = alphaOf(value);
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
  const options = ("options" in fill ? fill.options : fill) as {
    stops: readonly { position: number; color: unknown }[];
    shade?: unknown;
  };
  const stops = options.stops
    .map((stop) => ({
      offset: stop.position,
      color: colorPaintOf(stop.color, context),
    }))
    .filter((stop): stop is { offset: number; color: string } => stop.color != null);
  if (stops.length < 2) return undefined;
  const shade = options.shade as Record<string, unknown> | undefined;
  const shadeAngle = shade && "angle" in shade ? (shade.angle as number | undefined) : undefined;
  const path = shorthand?.path ?? (shade && "path" in shade ? (shade.path as string) : undefined);
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
    const alpha = alphaOf(fill.color);
    return {
      fill: color,
      ...(alpha != null && alpha < 1 ? { opacity: alpha } : {}),
    };
  }
  if (fill.type === "gradient") return { fill: gradientPaintOf(fill, width, height, context) };
  if (fill.type === "blip") {
    const url = pictureSrcOf(fill.data, fill.imageType);
    return url ? { fill: { type: "image", url, mode: "stretch" } } : {};
  }
  const pattern = fill as Extract<FillOptions, { type: "pattern" }>;
  const paint = patternTileSrcOf({
    ...pattern,
    ...(pattern.foregroundColor != null
      ? { foregroundColor: colorOf(pattern.foregroundColor, context) }
      : {}),
    ...(pattern.backgroundColor != null
      ? { backgroundColor: colorOf(pattern.backgroundColor, context) }
      : {}),
  });
  return { fill: { type: "image", url: paint, mode: "repeat", repeat: true } };
}
