// Chart color resolution: ChartSpaceOptions keeps DrawingML theme tokens, so
// DOCX/PPTX projections resolve them before the format-neutral chart painter
// reads fills, markers and label text.

import { transformColor, type ColorTransformOptions } from "./color-transform";

type ThemeColorToken = {
  value: string;
  transforms?: ColorTransformOptions;
};

interface ChartColorContext {
  themeColors?: Readonly<Record<string, string | undefined>>;
  colorMapping?: object;
}

const LONG_COLOR_TOKEN: Record<string, string> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
};

function colorOf(value: unknown, context: ChartColorContext): string | undefined {
  if (!isThemeColor(value)) return undefined;
  const token = value.value.replace("#", "");
  const hex = /^[0-9A-Fa-f]{6}$/.test(token) ? token.toUpperCase() : token;
  const key = LONG_COLOR_TOKEN[hex.toLowerCase()] ?? hex;
  const mapped = context.colorMapping
    ? Object.entries(context.colorMapping).find(([entryKey]) => entryKey === key)?.[1]
    : undefined;
  const slot = typeof mapped === "string" ? mapped : hex;
  const base = context.themeColors?.[slot];
  const resolved = transformColor(base ?? hex, value.transforms);
  return resolved.alpha >= 1
    ? resolved.color
    : `rgba(${hexToRgb(resolved.color)}, ${resolved.alpha})`;
}

function hexToRgb(hex: string): string {
  const value = hex.padStart(6, "0");
  return [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16)).join(", ");
}

function isThemeColor(value: unknown): value is ThemeColorToken {
  return (
    !!value &&
    typeof value === "object" &&
    typeof (value as { value?: unknown }).value === "string" &&
    Object.keys(value as { value?: unknown }).every(
      (key) => key === "value" || key === "transforms",
    )
  );
}

function resolve(value: unknown, context: ChartColorContext): unknown {
  if (Array.isArray(value)) return value.map((item) => resolve(item, context));
  if (!value || typeof value !== "object") return value;
  if (isThemeColor(value)) return colorOf(value, context) ?? value;
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, item]) => [
      key,
      resolve(item, context),
    ]),
  );
}

/** Resolve scheme tokens and color transforms in a chart's verbatim payload.
 *  Pure: the parsed chart stays available for edits. */
export function resolveChartColors<T>(chart: T, context: ChartColorContext): T {
  return resolve(chart, context) as T;
}
