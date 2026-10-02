// Chart color resolution: ChartSpaceOptions keeps DrawingML theme tokens, so
// the PPTX projection resolves them before the format-neutral chart painter
// reads fills, markers and label text.

import type { ColorMappingOptions } from "@office-open/core/theme";

import { transformColor } from "./color-transform";
import type { ThemeColors } from "./table-style";

interface ChartColorContext {
  themeColors?: ThemeColors;
  colorMapping?: ColorMappingOptions;
}

const LONG_COLOR_TOKEN: Record<string, string> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
};

function colorOf(value: unknown, context: ChartColorContext): string | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const token = record.value;
  if (typeof token !== "string") return undefined;
  const tokenValue = token.replace("#", "");
  const hex = /^[0-9A-Fa-f]{6}$/.test(tokenValue) ? tokenValue.toUpperCase() : tokenValue;
  const key = LONG_COLOR_TOKEN[hex.toLowerCase()] ?? hex;
  const slot = context.colorMapping?.[key as keyof ColorMappingOptions] ?? hex;
  const base = context.themeColors?.[slot as keyof ThemeColors];
  const resolved = transformColor(base ?? hex, record.transforms as never);
  return resolved.alpha >= 1
    ? resolved.color
    : `rgba(${hexToRgb(resolved.color)}, ${resolved.alpha})`;
}

function hexToRgb(hex: string): string {
  const value = hex.padStart(6, "0");
  return [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16)).join(", ");
}

function isThemeColor(value: unknown): value is Record<string, unknown> {
  return (
    !!value &&
    typeof value === "object" &&
    typeof (value as Record<string, unknown>).value === "string" &&
    Object.keys(value as Record<string, unknown>).every(
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
