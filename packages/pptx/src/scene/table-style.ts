// Table style resolution: a tblPr's styleId/tableStyle plus the six region
// flags → the normalized per-region fills/borders/text rules the projection
// applies. PowerPoint bakes its built-in styles into the app (the file
// carries only the GUID), so known GUIDs map to approximated families and
// everything unknown falls back to the default themed look; custom styles
// resolve from p:tblStyleLst or the inline a:tableStyle.

import type {
  TableStyleOptions,
  TableStyleRegion,
  ThemeableLineStyleOptions,
} from "@office-open/core/drawing";
import type { ColorSchemeOptions } from "@office-open/core/theme";
import type { TableOptions } from "@office-open/pptx";

/** One clrScheme slot's value spelling (hex string or a sysClr object). */
type SchemeColorValue = NonNullable<ColorSchemeOptions["accent1"]>;

/** One resolved region: the fill, text overrides, and border rules the
 *  projection layers onto cells that declare none of their own. */
export interface TableRegionRules {
  fill?: string;
  bold?: boolean;
  textColor?: string;
  borders?: Partial<
    Record<
      "top" | "right" | "bottom" | "left" | "insideH" | "insideV",
      { px: number; color?: string }
    >
  >;
}

export type TableStyleRegions = Partial<Record<TableStyleRegion, TableRegionRules>>;

export interface TableStyleFlags {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
}

export interface ResolvedTableStyle {
  flags: TableStyleFlags;
  regions: TableStyleRegions;
}

/** The theme color slots table styles resolve schemeClr tokens against. */
export type ThemeColors = Partial<
  Record<
    | "dark1"
    | "light1"
    | "dark2"
    | "light2"
    | "accent1"
    | "accent2"
    | "accent3"
    | "accent4"
    | "accent5"
    | "accent6",
    string
  >
>;

// Built-in GUIDs worth matching exactly; GUID comparison is case-insensitive.
const NO_STYLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";
const NO_STYLE_TABLE_GRID = "{5940675A-B579-460E-94D1-54222C63F5DA}";

const WHITE = "FFFFFF";
// Office 2013+ default accent1 — the fallback when the master carries no theme.
const DEFAULT_ACCENT = "4472C4";
// PowerPoint's inside rule for the Medium Style 2 family (~0.75 pt at 96 dpi).
const INSIDE_PX = 1;

/** Walk each channel toward white (fraction > 0) or black (< 0) — the
 *  light-accent banding look without the full color-transform model. */
function shadeOf(hex: string, fraction: number): string {
  const v = hex.replace("#", "").padStart(6, "0").slice(0, 6);
  const channel = (raw: string) => {
    const n = Number.parseInt(raw, 16);
    const out = fraction >= 0 ? n + (255 - n) * fraction : n * (1 + fraction);
    return Math.round(Math.min(255, Math.max(0, out)))
      .toString(16)
      .padStart(2, "0");
  };
  return `${channel(v.slice(0, 2))}${channel(v.slice(2, 4))}${channel(v.slice(4, 6))}`.toUpperCase();
}

function themeColorOf(value: SchemeColorValue | undefined, fallback: string): string {
  if (typeof value === "string") return value.replace("#", "").toUpperCase();
  return (value?.lastClr ?? fallback).replace("#", "").toUpperCase();
}

/** The Medium Style 2 look in accent terms: solid accent header/total rows
 *  with white bold text, white inside rules, light-accent banding. */
function mediumStyle2(theme: ThemeColors | undefined): TableStyleRegions {
  const accent = theme?.accent1?.replace("#", "").toUpperCase() ?? DEFAULT_ACCENT;
  return {
    wholeTbl: {
      borders: {
        insideH: { px: INSIDE_PX, color: WHITE },
        insideV: { px: INSIDE_PX, color: WHITE },
      },
    },
    firstRow: { fill: accent, bold: true, textColor: WHITE },
    lastRow: { fill: accent, bold: true, textColor: WHITE },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H: { fill: shadeOf(accent, 0.8) },
    band2H: { fill: shadeOf(accent, 0.6) },
  };
}

function gridRegions(): TableStyleRegions {
  const edge = { px: INSIDE_PX };
  return {
    wholeTbl: {
      borders: { top: edge, right: edge, bottom: edge, left: edge, insideH: edge, insideV: edge },
    },
  };
}

/** A raw fill choice (a:srgbClr or a:schemeClr with lum/tint children) →
 *  the hex the painter paints; unresolvable fills drop out. */
function fillOf(raw: string, theme: ThemeColors | undefined): string | undefined {
  const srgb = /srgbClr\s+val="([0-9a-fA-F]{6})"/.exec(raw);
  if (srgb) return srgb[1]!.toUpperCase();
  const scheme = /schemeClr\s+val="(\w+)"/.exec(raw);
  if (!scheme) return undefined;
  const base = themeColorOf(theme?.[scheme[1] as keyof ThemeColors], DEFAULT_ACCENT);
  // The common "lighter variation" spelling: lumMod 20% + lumOff 80% ≈ tint 0.8.
  if (/lumMod\s+val="20000"/.test(raw) && /lumOff\s+val="80000"/.test(raw))
    return shadeOf(base, 0.8);
  if (/lumMod\s+val="40000"/.test(raw) && /lumOff\s+val="60000"/.test(raw))
    return shadeOf(base, 0.6);
  const tint = /tint\s+val="(\d+)"/.exec(raw);
  if (tint) return shadeOf(base, Number(tint[1]) / 100000);
  const shade = /shade\s+val="(\d+)"/.exec(raw);
  if (shade) return shadeOf(base, -Number(shade[1]) / 100000);
  return base;
}

function lineOf(line: ThemeableLineStyleOptions, theme: ThemeColors | undefined) {
  // ThemeableLineStyleOptions.width is EMU (12700 = 1 pt); absent → 1 px.
  const px = line.width ? Math.max(1, Math.round(line.width / 9525)) : INSIDE_PX;
  return { px, ...(line.color ? { color: fillOf(line.color, theme) } : {}) };
}

function bordersOf(
  borders: Record<string, ThemeableLineStyleOptions | undefined> | undefined,
  theme: ThemeColors | undefined,
): TableRegionRules["borders"] {
  if (!borders) return undefined;
  const out: NonNullable<TableRegionRules["borders"]> = {};
  for (const key of ["top", "right", "bottom", "left", "insideH", "insideV"] as const) {
    const line = borders[key];
    if (line && !/noFill/.test(line.color ?? "")) out[key] = lineOf(line, theme);
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** An inline a:tableStyle (or a tblStyleLst entry) → region rules. */
export function regionsOf(
  style: TableStyleOptions,
  theme: ThemeColors | undefined,
): TableStyleRegions {
  const out: TableStyleRegions = {};
  for (const [region, part] of Object.entries(style.regions ?? {}) as [
    TableStyleRegion,
    (
      | {
          cell?: { borders?: Record<string, ThemeableLineStyleOptions | undefined>; fill?: string };
          text?: { bold?: "on" | "off" | "default"; color?: string };
        }
      | undefined
    ),
  ][]) {
    if (!part) continue;
    const rules: TableRegionRules = {};
    if (part.cell?.fill) rules.fill = fillOf(part.cell.fill, theme);
    const borders = bordersOf(part.cell?.borders, theme);
    if (borders) rules.borders = borders;
    if (part.text?.bold === "on") rules.bold = true;
    if (part.text?.color) rules.textColor = fillOf(part.text.color, theme);
    if (rules.fill || rules.bold || rules.textColor || rules.borders) out[region] = rules;
  }
  return out;
}

/** The flags a tblPr carries (attrs default off) plus the resolved regions:
 *  inline style wins, then a tblStyleLst match, then the GUID families. */
export function resolveTableStyle(
  table: TableOptions,
  theme: ThemeColors | undefined,
  custom: Record<string, TableStyleRegions> | undefined,
): ResolvedTableStyle {
  const flags: TableStyleFlags = {
    firstRow: table.firstRow === true,
    lastRow: table.lastRow === true,
    firstCol: table.firstCol === true,
    lastCol: table.lastCol === true,
    bandRow: table.bandRow === true,
    bandCol: table.bandCol === true,
  };
  if (table.tableStyle) return { flags, regions: regionsOf(table.tableStyle, theme) };
  const id = table.tableStyleId?.toUpperCase();
  if (id && custom) {
    const regions = custom[id];
    if (regions) return { flags, regions };
  }
  if (id === NO_STYLE_NO_GRID) return { flags, regions: {} };
  if (id === NO_STYLE_TABLE_GRID) return { flags, regions: gridRegions() };
  // No styleId and unknown built-ins all land on PowerPoint's default look.
  return { flags, regions: mediumStyle2(theme) };
}

/** The layered rules for one grid slot: wholeTbl base, then band, then the
 *  first/last overrides (PowerPoint's application order). */
export function regionRulesAt(
  style: ResolvedTableStyle,
  row: number,
  col: number,
  nRows: number,
  nCols: number,
): TableRegionRules {
  const out: TableRegionRules = {};
  const regions = style.regions;
  const apply = (key: TableStyleRegion) => {
    const part = regions[key];
    if (!part) return;
    if (part.fill !== undefined) out.fill = part.fill;
    if (part.bold !== undefined) out.bold = part.bold;
    if (part.textColor !== undefined) out.textColor = part.textColor;
    if (part.borders) out.borders = { ...out.borders, ...part.borders };
  };
  apply("wholeTbl");
  if (style.flags.bandRow) {
    const offset = style.flags.firstRow ? 1 : 0;
    apply((row - offset) % 2 === 0 ? "band1H" : "band2H");
  }
  if (style.flags.bandCol) {
    const offset = style.flags.firstCol ? 1 : 0;
    apply((col - offset) % 2 === 0 ? "band1V" : "band2V");
  }
  if (style.flags.firstRow && row === 0) apply("firstRow");
  if (style.flags.lastRow && row === nRows - 1) apply("lastRow");
  if (style.flags.firstCol && col === 0) apply("firstCol");
  if (style.flags.lastCol && col === nCols - 1) apply("lastCol");
  return out;
}
