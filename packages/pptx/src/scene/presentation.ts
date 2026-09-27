// The PPTX scene projection: PresentationOptions → flat slide-absolute drawing
// members the core painter paints. Batch 1 covers the vector surface —
// shapes, pictures, lines/connectors, groups, shape text and chart frames —
// through the same format-neutral extraction helpers the docx projection
// shares (@docen/core/geometry). Batch 2 adds graphic-frame tables (a:tbl →
// the frame-table painter's normalized payload). Batch 3 adds SmartArt
// (dgm:relIds → the painter's normalized fallback). Batch 4 adds media
// frames (posters plus stable player fallback). Table styles resolve through
// table-style.ts (tblPr flags, the themed default family, tblStyleLst
// entries); native media playback stays with its follow-up batch.

import { measureEmu, solidFillOf } from "@docen/core/geometry";
import { emuToPx, type LayoutDrawingMember } from "@docen/layout";
import type {
  ColorTransformOptions,
  FillOptions,
  SolidFillOptions,
} from "@office-open/core/drawing";
import type { StyleMatrixReferenceOptions } from "@office-open/core/drawing";
import type { ColorMappingOptions } from "@office-open/core/theme";
import { DEFAULT_COLOR_MAPPING } from "@office-open/core/theme";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

import { transformColor, type ResolvedColor } from "./color-transform";
import { IDENTITY } from "./geometry";
import { patternBackgroundOf } from "./pattern-background";
import { pictureSrcOf } from "./pictures";
import { regionsOf, type ThemeColors } from "./table-style";
import { childMembers } from "./walk";

/** One projected presentation: the slide size plus each slide's members. */
export interface ProjectedPresentation {
  /** Slide width in px (96 dpi; EMU ÷ 9525). */
  widthPx: number;
  /** Slide height in px. */
  heightPx: number;
  slides: ProjectedSlide[];
}

/** One projected slide. */
export interface ProjectedSlide {
  /** The slide's background paint; absent → the painter's page default. */
  background?: ProjectedSlideBackground;
  /** Slide-absolute drawing members, paint order = document order. */
  members: LayoutDrawingMember[];
}

/** The background forms the painter renders: a solid hex, a gradient's
 *  stops with its linear angle (degrees, OOXML clockwise-from-east) or radial
 *  path, or a picture fill's data URL. */
export type ProjectedSlideBackground =
  | { kind: "solid"; color: string }
  | {
      kind: "gradient";
      stops: { color: string; position: number }[];
      angle?: number;
      path?: "shape" | "circle" | "rect";
    }
  | { kind: "image"; src: string };

/** Project a parsed presentation into the paintable shape: sizes are resolved
 *  to px and every slide's children flatten into members. Pure — the input is
 *  not mutated, so callers may re-project on demand. */
export function projectPresentation(pres: PresentationOptions): ProjectedPresentation {
  const { widthPx, heightPx } = slideSizePx(pres.size);
  const context = tableContextOf(pres);
  const size = { widthPx, heightPx };
  return {
    widthPx,
    heightPx,
    slides: (pres.slides ?? []).map((slide, index) =>
      projectSlide(slide, index + 1, pres.slides?.length, context, pres.masters, size),
    ),
  };
}

/** The theme bits the projection resolves against: scheme colors, the bg
 * fill style list, and the token → slot color map (the master's clrMap). */
interface ThemeContext {
  themeColors?: ThemeColors;
  tableStyles?: Record<string, ReturnType<typeof regionsOf>>;
  backgroundFillStyles: FillOptions[];
  colorMapping: ColorMappingOptions;
}

/** The theme/table-style context every slide's projection shares: the first
 * master's color scheme plus the tblStyleLst entries keyed by GUID. */
function tableContextOf(pres: PresentationOptions): ThemeContext {
  const master = pres.masters?.[0];
  const scheme = master?.theme?.colorScheme;
  const keys = [
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
  ] as const;
  const themeColors: ThemeColors | undefined = scheme
    ? Object.fromEntries(
        keys
          .map((key) => [key, themeColorOf(scheme[key])] as const)
          .filter((entry): entry is [(typeof keys)[number], string] => entry[1] !== undefined),
      )
    : undefined;
  const tableStyles = pres.tableStyles?.styles?.length
    ? Object.fromEntries(
        pres.tableStyles.styles.map((style) => [
          style.styleId.toUpperCase(),
          regionsOf(style, themeColors),
        ]),
      )
    : undefined;
  return {
    themeColors,
    tableStyles,
    backgroundFillStyles: master?.theme?.formatScheme?.backgroundFillStyles ?? [],
    colorMapping: { ...DEFAULT_COLOR_MAPPING, ...master?.colorMapping },
  };
}

/** One theme slot's hex (a string value or a sysClr's lastClr). */
function themeColorOf(value: string | { lastClr?: string } | undefined): string | undefined {
  if (value === undefined) return undefined;
  const hex = typeof value === "string" ? value : value.lastClr;
  return hex ? hex.replace("#", "").toUpperCase() : undefined;
}

const NAMED_SLIDE_PX = {
  "16:9": { width: 1280, height: 720 },
  "4:3": { width: 960, height: 720 },
} as const;

/** The named slide classes at 96 dpi; an explicit size resolves its EMU.
 *  Absent → 16:9 (PowerPoint's default). */
function slideSizePx(size: PresentationOptions["size"]): { widthPx: number; heightPx: number } {
  if (size === "4:3")
    return { widthPx: NAMED_SLIDE_PX["4:3"].width, heightPx: NAMED_SLIDE_PX["4:3"].height };
  if (size === "16:9" || size === undefined)
    return { widthPx: NAMED_SLIDE_PX["16:9"].width, heightPx: NAMED_SLIDE_PX["16:9"].height };
  return {
    widthPx: emuToPx(measureEmu(size.width) ?? 0),
    heightPx: emuToPx(measureEmu(size.height) ?? 0),
  };
}

function projectSlide(
  slide: SlideOptions,
  slideNumber: number,
  slideCount: number | undefined,
  context: ThemeContext,
  masters: PresentationOptions["masters"],
  size: { widthPx: number; heightPx: number },
): ProjectedSlide {
  const background = slideBackgroundOf(slide, masters, context, size);
  return {
    ...(background ? { background } : {}),
    members: childMembers(slide.children ?? [], IDENTITY, [], {
      slideNumber,
      slideCount,
      now: new Date(),
      ...context,
    }),
  };
}

/** The slide's background paint: its own fill/reference, or the master's
 * (slides inherit the master background when they declare none). */
function slideBackgroundOf(
  slide: SlideOptions,
  masters: PresentationOptions["masters"],
  context: ThemeContext,
  size: { widthPx: number; heightPx: number },
): ProjectedSlideBackground | undefined {
  const background = slide.background ?? masters?.[0]?.background;
  if (!background) return undefined;
  if (background.fill)
    return backgroundOf(
      resolveFillColors(background.fill, undefined, context),
      size.widthPx,
      size.heightPx,
    );
  if (background.reference)
    return bgRefFill(background.reference, context, size.widthPx, size.heightPx);
  return undefined;
}

/** A p:bgRef style-matrix reference → the theme's background fill style at
 * the index (bgFillStyleLst addresses from 1001), its phClr placeholders
 * replaced with the reference's color. */
function bgRefFill(
  reference: StyleMatrixReferenceOptions,
  context: ThemeContext,
  widthPx: number,
  heightPx: number,
): ProjectedSlideBackground | undefined {
  const style =
    context.backgroundFillStyles[
      reference.index >= 1001 ? reference.index - 1001 : reference.index
    ];
  if (!style) return undefined;
  const phClr = styleColorOf(reference.color, context, undefined);
  return backgroundOf(resolveFillColors(style, phClr, context), widthPx, heightPx);
}

/** One EG_ColorChoice → hex: hex strings pass, scheme tokens resolve through
 * the master's color map into the theme's scheme, phClr takes the style
 * reference's color, then evaluates its color-transform list. */
function resolvedStyleColorOf(
  color: unknown,
  context: ThemeContext,
  phClr: string | undefined,
): ResolvedColor | undefined {
  if (typeof color === "string") {
    const hex = color.replace("#", "").toUpperCase();
    return /^[0-9A-F]{6}$/.test(hex) ? { color: hex, alpha: 1 } : undefined;
  }
  if (color && typeof color === "object" && "value" in color && typeof color.value === "string") {
    const token = color.value;
    if (token === "phClr") return phClr ? transformColor(phClr, transformsOf(color)) : undefined;
    if (/^[0-9A-F]{6}$/i.test(token)) return transformColor(token, transformsOf(color));
    // The raw clrMap tokens spell the first four slots short (bg1/tx1/bg2/tx2);
    // the parsed color map keys them long (background1/text1/…).
    const key =
      { bg1: "background1", tx1: "text1", bg2: "background2", tx2: "text2" }[token] ?? token;
    const slot = (context.colorMapping as unknown as Record<string, string>)[key] ?? token;
    const hex = context.themeColors?.[slot as keyof ThemeColors];
    return hex ? transformColor(hex, transformsOf(color)) : undefined;
  }
  return undefined;
}

function styleColorOf(
  color: unknown,
  context: ThemeContext,
  phClr: string | undefined,
): string | undefined {
  return resolvedStyleColorOf(color, context, phClr)?.color;
}

function stylePaintOf(
  color: unknown,
  context: ThemeContext,
  phClr: string | undefined,
): string | undefined {
  const resolved = resolvedStyleColorOf(color, context, phClr);
  if (!resolved) return undefined;
  if (resolved.alpha >= 1) return resolved.color;
  const raw = Number.parseInt(resolved.color, 16);
  return `rgba(${(raw >> 16) & 255}, ${(raw >> 8) & 255}, ${raw & 255}, ${resolved.alpha})`;
}

/** Replace every phClr/scheme color in a style-matrix fill with its resolved
 * hex (a shallow structural walk — the fill shapes the projection reads). */
function resolveFillColors(
  fill: FillOptions,
  phClr: string | undefined,
  context: ThemeContext,
): FillOptions {
  if (typeof fill === "string") return fill;
  if (fill.type === "solid") {
    return {
      ...fill,
      color: colorChoiceOf(fill.color, context, phClr) as never,
    };
  }
  if (fill.type === "gradient") {
    const resolveStops = <T extends { position: number; color: unknown }>(
      stops: readonly T[],
    ): T[] =>
      stops.map(
        (stop) => ({ ...stop, color: stylePaintOf(stop.color, context, phClr) ?? stop.color }) as T,
      );
    if ("options" in fill) {
      return {
        type: "gradient",
        options: { ...fill.options, stops: resolveStops(fill.options.stops) },
      };
    }
    return { ...fill, stops: resolveStops(fill.stops) } as FillOptions;
  }
  if (fill.type === "pattern") {
    return {
      ...fill,
      foregroundColor: colorChoiceOf(fill.foregroundColor, context, phClr) as never,
      backgroundColor: colorChoiceOf(fill.backgroundColor, context, phClr) as never,
    };
  }
  return fill;
}

/** A SolidFillOptions wrapper or raw EG_ColorChoice → the resolver. */
function colorChoiceOf(
  value: unknown,
  context: ThemeContext,
  phClr: string | undefined,
): string | undefined {
  if (value && typeof value === "object" && "type" in value && value.type === "solid") {
    return stylePaintOf((value as unknown as { color: unknown }).color, context, phClr);
  }
  return stylePaintOf(value, context, phClr);
}

function transformsOf(color: unknown): ColorTransformOptions | undefined {
  if (!color || typeof color !== "object" || !("transforms" in color)) return undefined;
  const transforms = (color as { transforms?: unknown }).transforms;
  return transforms && typeof transforms === "object"
    ? (transforms as ColorTransformOptions)
    : undefined;
}

/** A p:bg fill → the projected paint; exotic fills keep the painter's default. */
function backgroundOf(
  fill: FillOptions,
  widthPx: number,
  heightPx: number,
): ProjectedSlideBackground | undefined {
  if (typeof fill === "string") {
    const color = backgroundColorOf(fill);
    return color ? { kind: "solid", color } : undefined;
  }
  if (fill.type === "solid") {
    const color = backgroundColorOf(fill.color);
    return color ? { kind: "solid", color } : undefined;
  }
  if (fill.type === "gradient") {
    const opts = "options" in fill ? fill.options : fill;
    // Stop colors follow the core color shape (a hex string or a
    // {value} record) — not the docen solid-fill wrapper solidFillOf reads.
    const stopColor = (c: unknown) => {
      if (typeof c === "string") {
        const paint = c.trim();
        if (/^rgba\(/i.test(paint)) return paint;
        return paint.replace("#", "").toUpperCase();
      }
      return c && typeof c === "object" && "value" in c && typeof c.value === "string"
        ? c.value.replace("#", "").toUpperCase()
        : undefined;
    };
    const stops = opts.stops
      .map((stop) => ({ color: stopColor(stop.color) ?? "", position: stop.position }))
      .filter((stop) => stop.color !== "");
    if (stops.length < 2) return undefined;
    // The shorthand carries angle/path at the top level; the full options
    // form nests them inside shade (linear vs path).
    const shorthand = "options" in fill ? undefined : fill;
    const shade = opts && typeof opts === "object" && "shade" in opts ? opts.shade : undefined;
    const angle = shorthand?.angle ?? (shade && "angle" in shade ? shade.angle : undefined);
    const path = shorthand?.path ?? (shade && "path" in shade ? shade.path : undefined);
    return {
      kind: "gradient",
      stops,
      ...(angle !== undefined ? { angle } : {}),
      ...(path ? { path } : {}),
    };
  }
  if (fill.type === "blip") {
    const src = pictureSrcOf(fill.data, fill.imageType);
    return src ? { kind: "image", src } : undefined;
  }
  if (fill.type === "pattern") {
    const src = patternBackgroundOf(fill, widthPx, heightPx);
    return src ? { kind: "image", src } : undefined;
  }
  return undefined;
}

/** A resolved fill color may now carry alpha as CSS rgba(); hex stays the
 *  painter's canonical RRGGBB form. */
function backgroundColorOf(value: string | SolidFillOptions): string | undefined {
  if (typeof value === "string") {
    if (/^rgba\(/i.test(value)) return value;
  }
  return solidFillOf(value);
}
