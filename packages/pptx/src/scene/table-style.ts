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
import type { ColorTransformOptions } from "@office-open/core/drawing";
import type { ColorSchemeOptions } from "@office-open/core/theme";
import type { TableOptions } from "@office-open/pptx";

import { transformColor } from "./color-transform";

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

/** The theme color slots the projection resolves schemeClr tokens against —
 *  the resolved hex per the office-open color scheme's own slot set. */
export type ThemeColors = Partial<Record<Exclude<keyof ColorSchemeOptions, "name">, string>>;

/** The scheme slots the projection reads, typed against the office-open
 *  color scheme so a renamed slot breaks compilation here, not at runtime. */
export const SCHEME_SLOTS = [
  "dark1",
  "light1",
  "dark2",
  "light2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const satisfies readonly (keyof ThemeColors)[];

const WHITE = "FFFFFF";
const BLACK = "000000";
// Office 2013+ default accent1 — the fallback when the master carries no theme.
const DEFAULT_ACCENT = "4472C4";
// PowerPoint's inside rule for the Medium Style 2 family (~0.75 pt at 96 dpi).
const INSIDE_PX = 1;

type AccentNumber = 1 | 2 | 3 | 4 | 5 | 6;
type BuiltinFamily =
  | "noStyle"
  | "tableGrid"
  | "themed1"
  | "themed2"
  | "light1"
  | "light2"
  | "light3"
  | "medium1"
  | "medium2"
  | "medium3"
  | "medium4"
  | "dark1"
  | "dark2";

interface BuiltinStyle {
  family: BuiltinFamily;
  accent?: AccentNumber;
  secondaryAccent?: AccentNumber;
}

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

/** Resolve the DrawingML color token, accepting the short spellings built-in
 *  table styles use (`dk1`, `tx1`, `lt1`, `bg1`). */
function schemeTokenOf(token: string, theme: ThemeColors | undefined, fallback: string): string {
  const normalized = token
    .replace(/^dk1$/i, "dark1")
    .replace(/^lt1$/i, "light1")
    .replace(/^tx1$/i, "dark1")
    .replace(/^bg1$/i, "light1")
    .replace(/^(dk2|tx2)$/i, "dark2")
    .replace(/^(lt2|bg2)$/i, "light2");
  return themeColorOf(theme?.[normalized as keyof ThemeColors], fallback);
}

function accentOf(theme: ThemeColors | undefined, accent?: AccentNumber): string {
  if (!accent) return schemeTokenOf("dk1", theme, BLACK);
  return themeColorOf(theme?.[`accent${accent}` as keyof ThemeColors], DEFAULT_ACCENT);
}

function borderOf(
  color?: string,
  px = INSIDE_PX,
): NonNullable<TableRegionRules["borders"]>[keyof NonNullable<TableRegionRules["borders"]>] {
  return { px, ...(color ? { color } : {}) };
}

function bordersOfKeys(
  color: string | undefined,
  px: number,
  keys: readonly ("top" | "right" | "bottom" | "left" | "insideH" | "insideV")[],
): NonNullable<TableRegionRules["borders"]> {
  return Object.fromEntries(keys.map((key) => [key, borderOf(color, px)]));
}

/** Approximate a translucent solid fill after compositing on a white slide. */
function washOf(hex: string, alpha: number): string {
  return shadeOf(hex, alpha);
}

/** Themed Style 1: an outlined table with translucent accent bands. */
function themed1(theme: ThemeColors | undefined, accent: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: {
      borders: bordersOfKeys(color, 1, ["top", "right", "bottom", "left", "insideH", "insideV"]),
    },
    firstRow: { fill: color, bold: true, textColor: WHITE, borders: { bottom: borderOf(WHITE) } },
    lastRow: { bold: true, borders: { top: borderOf(color) } },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H: { fill: washOf(color, 0.6) },
    band1V: { fill: washOf(color, 0.6), borders: bordersOfKeys(color, 1, ["top", "bottom"]) },
  };
}

/** Themed Style 2: an accent outline, bold separators, and subtle white bands. */
function themed2(theme: ThemeColors | undefined, accent: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: { borders: bordersOfKeys(color, 1, ["top", "right", "bottom", "left"]) },
    firstRow: { bold: true, borders: { bottom: borderOf(WHITE) } },
    lastRow: { bold: true, borders: { top: borderOf(WHITE) } },
    firstCol: { bold: true, borders: { right: borderOf(WHITE) } },
    lastCol: { bold: true, borders: { left: borderOf(WHITE) } },
    band1H: { fill: washOf(WHITE, 0.2) },
    band1V: { fill: washOf(WHITE, 0.2) },
  };
}

/** The three light families differ mainly in outline weight and banding. */
function lightStyle(
  family: "light1" | "light2" | "light3",
  theme: ThemeColors | undefined,
  accent?: AccentNumber,
): TableStyleRegions {
  const color = accentOf(theme, accent);
  const whole =
    family === "light1"
      ? { borders: { top: borderOf(color), bottom: borderOf(color) } }
      : family === "light2"
        ? { borders: bordersOfKeys(color, 1, ["top", "right", "bottom", "left"]) }
        : {
            borders: bordersOfKeys(color, 1, [
              "top",
              "right",
              "bottom",
              "left",
              "insideH",
              "insideV",
            ]),
          };
  return {
    wholeTbl: whole,
    firstRow: {
      bold: true,
      borders:
        family === "light2" ? undefined : { bottom: borderOf(color, family === "light1" ? 1 : 2) },
    },
    lastRow: { bold: true, borders: { top: borderOf(color, family === "light1" ? 1 : 4) } },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H:
      family === "light2"
        ? { borders: { top: borderOf(color), bottom: borderOf(color) } }
        : { fill: washOf(color, 0.8) },
    band1V:
      family === "light2"
        ? { borders: { left: borderOf(color), right: borderOf(color) } }
        : family === "light1"
          ? { fill: washOf(color, 0.8) }
          : undefined,
  };
}

function medium1(theme: ThemeColors | undefined, accent?: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: {
      fill: WHITE,
      borders: bordersOfKeys(color, 1, ["top", "right", "bottom", "left", "insideH"]),
    },
    firstRow: { fill: color, bold: true, textColor: WHITE },
    lastRow: { fill: WHITE, bold: true, borders: { top: borderOf(color, 4) } },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H: { fill: shadeOf(color, 0.8) },
    band1V: { fill: shadeOf(color, 0.8) },
  };
}

/** Medium Style 2 is PowerPoint's default; all dk1 slots move to the accent. */
function medium2(theme: ThemeColors | undefined, accent?: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: {
      fill: shadeOf(color, 0.2),
      borders: bordersOfKeys(WHITE, 1, ["top", "right", "bottom", "left", "insideH", "insideV"]),
    },
    firstRow: {
      fill: color,
      bold: true,
      textColor: WHITE,
      borders: { bottom: borderOf(WHITE, 3) },
    },
    lastRow: { fill: color, bold: true, textColor: WHITE, borders: { top: borderOf(WHITE, 3) } },
    firstCol: { fill: color, bold: true, textColor: WHITE },
    lastCol: { fill: color, bold: true, textColor: WHITE },
    band1H: { fill: shadeOf(color, 0.4) },
    band2H: undefined,
    band1V: { fill: shadeOf(color, 0.4) },
    band2V: undefined,
  };
}

function medium3(theme: ThemeColors | undefined, accent?: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: { fill: WHITE, borders: { top: borderOf(color, 2), bottom: borderOf(color, 2) } },
    firstRow: {
      fill: color,
      bold: true,
      textColor: WHITE,
      borders: { bottom: borderOf(color, 2) },
    },
    lastRow: { fill: WHITE, bold: true, borders: { top: borderOf(color, 4) } },
    firstCol: { fill: color, bold: true, textColor: WHITE },
    lastCol: { fill: color, bold: true, textColor: WHITE },
    band1H: { fill: shadeOf(BLACK, 0.2) },
    band1V: { fill: shadeOf(BLACK, 0.2) },
  };
}

function medium4(theme: ThemeColors | undefined, accent?: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: {
      fill: shadeOf(color, 0.2),
      borders: bordersOfKeys(color, 1, ["top", "right", "bottom", "left", "insideH", "insideV"]),
    },
    firstRow: { fill: shadeOf(color, 0.2), bold: true, textColor: WHITE },
    lastRow: {
      fill: shadeOf(color, 0.2),
      bold: true,
      textColor: WHITE,
      borders: { top: borderOf(color, 2) },
    },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H: { fill: shadeOf(color, 0.4) },
    band1V: { fill: shadeOf(color, 0.4) },
  };
}

function dark1(theme: ThemeColors | undefined, accent?: AccentNumber): TableStyleRegions {
  const color = accentOf(theme, accent);
  return {
    wholeTbl: { fill: shadeOf(color, 0.2) },
    firstRow: {
      fill: color,
      bold: true,
      textColor: WHITE,
      borders: { bottom: borderOf(WHITE, 2) },
    },
    lastRow: {
      fill: shadeOf(color, 0.6),
      bold: true,
      textColor: WHITE,
      borders: { top: borderOf(WHITE, 2) },
    },
    firstCol: {
      fill: shadeOf(color, 0.6),
      bold: true,
      textColor: WHITE,
      borders: { right: borderOf(WHITE, 2) },
    },
    lastCol: {
      fill: shadeOf(color, 0.6),
      bold: true,
      textColor: WHITE,
      borders: { left: borderOf(WHITE, 2) },
    },
    band1H: { fill: shadeOf(color, 0.4) },
    band1V: { fill: shadeOf(color, 0.4) },
  };
}

/** Dark Style 2's pair variants use one accent for the header and another for
 *  the body; the total row's top rule deliberately stays the original dark. */
function dark2(
  theme: ThemeColors | undefined,
  primary?: AccentNumber,
  secondary?: AccentNumber,
): TableStyleRegions {
  const headerColor = accentOf(theme, primary);
  const bodyColor = accentOf(theme, secondary);
  return {
    wholeTbl: { fill: shadeOf(bodyColor, 0.2) },
    firstRow: {
      fill: headerColor,
      bold: true,
      textColor: WHITE,
      borders: { bottom: borderOf(shadeOf(bodyColor, 0.2), 2) },
    },
    lastRow: {
      fill: shadeOf(bodyColor, 0.2),
      bold: true,
      textColor: WHITE,
      borders: { top: borderOf(BLACK, 4) },
    },
    firstCol: { bold: true },
    lastCol: { bold: true },
    band1H: { fill: shadeOf(bodyColor, 0.4) },
    band1V: { fill: shadeOf(bodyColor, 0.4) },
  };
}

function builtinRegions(style: BuiltinStyle, theme: ThemeColors | undefined): TableStyleRegions {
  switch (style.family) {
    case "noStyle":
      return {};
    case "tableGrid":
      return gridRegions();
    case "themed1":
      return themed1(theme, style.accent ?? 1);
    case "themed2":
      return themed2(theme, style.accent ?? 1);
    case "light1":
    case "light2":
    case "light3":
      return lightStyle(style.family, theme, style.accent);
    case "medium1":
      return medium1(theme, style.accent);
    case "medium2":
      return medium2(theme, style.accent);
    case "medium3":
      return medium3(theme, style.accent);
    case "medium4":
      return medium4(theme, style.accent);
    case "dark1":
      return dark1(theme, style.accent);
    case "dark2":
      return dark2(theme, style.accent, style.secondaryAccent);
  }
}

const withAccent = (family: BuiltinFamily, ids: readonly string[]): [string, BuiltinStyle][] =>
  ids.map((id, index) => [id, { family, accent: (index + 1) as AccentNumber }]);

const withBaseAndAccents = (
  family: BuiltinFamily,
  ids: readonly string[],
): [string, BuiltinStyle][] =>
  ids.map((id, index) => [
    id,
    index === 0 ? { family } : { family, accent: index as AccentNumber },
  ]);

/** MS-OE376's 74 built-in table styles, keyed by GUID. Base styles omit the
 *  accent replacement; arrays below are ordered Accent 1 through Accent 6. */
const BUILT_IN_STYLES: Record<string, BuiltinStyle> = Object.fromEntries([
  ["{2D5ABB26-0587-4C30-8999-92F81FD0307C}", { family: "noStyle" as const }],
  ["{5940675A-B579-460E-94D1-54222C63F5DA}", { family: "tableGrid" as const }],
  ...withAccent("themed1", [
    "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}",
    "{284E427A-3D55-4303-BF80-6455036E1DE7}",
    "{69C7853C-536D-4A76-A0AE-DD22124D55A5}",
    "{775DCB02-9BB8-47FD-8907-85C794F793BA}",
    "{35758FB7-9AC5-4552-8A53-C91805E547FA}",
    "{08FB837D-C827-4EFA-A057-4D05807E0F7C}",
  ]),
  ...withAccent("themed2", [
    "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}",
    "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}",
    "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}",
    "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}",
    "{327F97BB-C833-4FB7-BDE5-3F7075034690}",
    "{638B1855-1B75-4FBE-930C-398BA8C253C6}",
  ]),
  ...withBaseAndAccents("light1", [
    "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
    "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
    "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
    "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
    "{D27102A9-8310-4765-A935-A1911B00CA55}",
    "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}",
    "{68D230F3-CF80-4859-8CE7-A43EE81993B5}",
  ]),
  ...withBaseAndAccents("light2", [
    "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}",
    "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
    "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}",
    "{F2DE63D5-997A-4646-A377-4702673A728D}",
    "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}",
    "{5A111915-BE36-4E01-A7E5-04B1672EAD32}",
    "{912C8C85-51F0-491E-9774-3900AFEF0FD7}",
  ]),
  ...withBaseAndAccents("light3", [
    "{616DA210-FB5B-4158-B5E0-FEB733F419BA}",
    "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}",
    "{5DA37D80-6434-44D0-A028-1B22A696006F}",
    "{8799B23B-EC83-4686-B30A-512413B5E67A}",
    "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
    "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}",
    "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}",
  ]),
  ...withBaseAndAccents("medium1", [
    "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}",
    "{B301B821-A1FF-4177-AEE7-76D212191A09}",
    "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}",
    "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}",
    "{1E171933-4619-4E11-9A3F-F7608DF75F80}",
    "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}",
    "{10A1B5D5-9B99-4C35-A422-299274C87663}",
  ]),
  ...withBaseAndAccents("medium2", [
    "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
    "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
    "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
    "{00A15C55-8517-42AA-B614-E9B94910E393}",
    "{7DF18680-E054-41AD-8BC1-D1AEF772440D}",
    "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
  ]),
  ...withBaseAndAccents("medium3", [
    "{8EC20E35-A176-4012-BC5E-935CFFF8708E}",
    "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}",
    "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}",
    "{EB344D84-9AFB-497E-A393-DC336BA19D2E}",
    "{EB9631B5-78F2-41C9-869B-9F39066F8104}",
    "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}",
    "{2A488322-F2BA-4B5B-9748-0D474271808F}",
  ]),
  ...withBaseAndAccents("medium4", [
    "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
    "{69CF1AB2-1976-4502-BF36-3FF5EA218861}",
    "{8A107856-5554-42FB-B03E-39F5DBC370BA}",
    "{0505E3EF-67EA-436B-97B2-0124C06EBD24}",
    "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}",
    "{22838BEF-8BB2-4498-84A7-C5851F593DF1}",
    "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}",
  ]),
  ...withBaseAndAccents("dark1", [
    "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
    "{125E5076-3810-47DD-B79F-674D7AD40C01}",
    "{37CE84F3-28C3-443E-9E96-99CF82512B78}",
    "{D03447BB-5D67-496B-8E87-E561075AD55C}",
    "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}",
    "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}",
    "{AF606853-7671-496A-8E4F-DF71F8EC918B}",
  ]),
  ["{5202B0CA-FC54-4496-8BCA-5EF66A818D29}", { family: "dark2" as const }],
  [
    "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}",
    { family: "dark2" as const, accent: 2, secondaryAccent: 1 },
  ],
  [
    "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}",
    { family: "dark2" as const, accent: 4, secondaryAccent: 3 },
  ],
  [
    "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}",
    { family: "dark2" as const, accent: 6, secondaryAccent: 5 },
  ],
]);

function gridRegions(): TableStyleRegions {
  const edge = borderOf(BLACK);
  return {
    wholeTbl: {
      borders: { top: edge, right: edge, bottom: edge, left: edge, insideH: edge, insideV: edge },
    },
  };
}

/** A raw fill choice (a:srgbClr or a:schemeClr with transform children) →
 *  the hex the painter paints; unresolvable fills drop out. */
function fillOf(raw: string, theme: ThemeColors | undefined): string | undefined {
  const srgb = /srgbClr\s+val="([0-9a-fA-F]{6})"/.exec(raw);
  if (srgb) return srgb[1]!.toUpperCase();
  const scheme = /schemeClr\s+val="(\w+)"/.exec(raw);
  if (!scheme) return undefined;
  const base = schemeTokenOf(scheme[1]!, theme, DEFAULT_ACCENT);
  const transforms: Record<string, number | boolean> = {};
  const child = /<a:(\w+)(?:\s+val="(-?\d+)")?\s*\/>/g;
  for (const [, name, value] of raw.matchAll(child)) {
    if (!name || name === "srgbClr" || name === "schemeClr") continue;
    transforms[name as keyof ColorTransformOptions] = value == null ? true : Number(value) / 1000;
  }
  return transformColor(base, transforms as ColorTransformOptions).color;
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
  const builtin = id ? BUILT_IN_STYLES[id] : undefined;
  if (builtin) return { flags, regions: builtinRegions(builtin, theme) };
  // An unstyled a:tbl is PowerPoint's plain black grid: no theme fill, no
  // region overrides. The Medium Style 2 fallback painted bare office-open
  // tables as if a style reference had been omitted but still applied.
  return { flags, regions: gridRegions() };
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
