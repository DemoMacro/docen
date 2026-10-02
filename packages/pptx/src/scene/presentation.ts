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
import { browserFontMetrics, emuToPx } from "@docen/layout";
import type {
  ColorTransformOptions,
  FillOptions,
  GradientStopOptions,
  SolidFillOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
  TextListStyleOptions,
  TextStylesOptions,
} from "@office-open/core/drawing";
import type { StyleMatrixReferenceOptions } from "@office-open/core/drawing";
import type { ColorMappingOptions, ThemeOptions } from "@office-open/core/theme";
import { DEFAULT_COLOR_MAPPING } from "@office-open/core/theme";
import {
  resolvePlaceholder,
  type LayoutDefinition,
  type MasterDefinition,
  type PresentationOptions,
  type SlideChild,
  type SlideOptions,
} from "@office-open/pptx";

import { tilePaintOf } from "./blip-tile";
import { transformColor, type ResolvedColor } from "./color-transform";
import { IDENTITY } from "./geometry";
import { patternBackgroundOf } from "./pattern-background";
import { pictureSrcOf } from "./pictures";
import { regionsOf, SCHEME_SLOTS, type ThemeColors } from "./table-style";
import type { ThemeFonts } from "./text";
import { childMembers, type ProjectedSlideMember } from "./walk";

/** One projected presentation: the slide size plus each slide's members. */
export interface ProjectedPresentation {
  /** Slide width in px (96 dpi; EMU ÷ 9525). */
  widthPx: number;
  /** Slide height in px. */
  heightPx: number;
  slides: ProjectedSlide[];
}

type TextLevelStyle = NonNullable<TextListStyleOptions["levels"]>[number];

/** One projected slide. */
export interface ProjectedSlide {
  /** The slide's background paint; absent → the painter's page default. */
  background?: ProjectedSlideBackground;
  /** The slide's children with placeholder geometry and facets resolved for
   *  selection and hit testing; layout templates are excluded. */
  children: SlideChild[];
  /** Slide-absolute drawing members, paint order = document order. */
  members: ProjectedSlideMember[];
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
  | {
      kind: "image";
      src: string;
      tile?: {
        scale?: { x: number; y: number };
        offset?: { x: number; y: number };
        align?:
          | "top-left"
          | "top"
          | "top-right"
          | "left"
          | "center"
          | "right"
          | "bottom-left"
          | "bottom"
          | "bottom-right";
      };
    };

/** Project a parsed presentation into the paintable shape: sizes are resolved
 *  to px and every slide's children flatten into members. Pure — the input is
 *  not mutated, so callers may re-project on demand. */
export function projectPresentation(pres: PresentationOptions): ProjectedPresentation {
  const { widthPx, heightPx } = slideSizePx(pres.size);
  const size = { widthPx, heightPx };
  return {
    widthPx,
    heightPx,
    slides: (pres.slides ?? []).map((slide, index) =>
      projectSlide(pres, slide, index + 1, pres.slides?.length, pres.masters, size),
    ),
  };
}
export type { ProjectedSlideMember };

/** The theme bits the projection resolves against: scheme colors, the bg
 * fill style list, and the token → slot color map (the master's clrMap). */
interface ThemeContext {
  themeColors?: ThemeColors;
  themeFonts?: ThemeFonts;
  tableStyles?: Record<string, ReturnType<typeof regionsOf>>;
  backgroundFillStyles: FillOptions[];
  lineStyles?: NonNullable<ThemeOptions["formatScheme"]>["lineStyles"];
  colorMapping: ColorMappingOptions;
  textStyles?: TextStylesOptions;
}

/** The theme/table-style context for a slide's owning master: its color and
 * font schemes plus the tblStyleLst entries keyed by GUID. */
function tableContextOf(
  pres: PresentationOptions,
  masterName: SlideOptions["master"],
  masters: PresentationOptions["masters"],
): ThemeContext {
  const master = masters?.find((candidate) => candidate.name === masterName) ?? masters?.[0];
  const scheme = master?.theme?.colorScheme;
  const themeColors: ThemeColors | undefined = scheme
    ? Object.fromEntries(
        SCHEME_SLOTS.map((key) => [key, themeColorOf(scheme[key])] as const).filter(
          (entry): entry is [(typeof SCHEME_SLOTS)[number], string] => entry[1] !== undefined,
        ),
      )
    : undefined;
  const themeFonts = themeFontsOf(master?.theme?.fontScheme);
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
    themeFonts,
    tableStyles,
    backgroundFillStyles: master?.theme?.formatScheme?.backgroundFillStyles ?? [],
    lineStyles: master?.theme?.formatScheme?.lineStyles,
    colorMapping: { ...DEFAULT_COLOR_MAPPING, ...master?.colorMapping },
    textStyles: master?.textStyles,
  };
}

/** A font scheme's non-empty faces; the Office defaults leave ea/cs empty,
 * and an empty face must fall through to the painter's default family. */
function themeFontsOf(scheme: ThemeOptions["fontScheme"]): ThemeFonts | undefined {
  if (!scheme) return undefined;
  const faces = (collection: typeof scheme.majorFont): ThemeFonts["minor"] => {
    if (!collection) return undefined;
    const face = (font: { typeface?: string } | undefined): string | undefined =>
      font?.typeface || undefined;
    const out = {
      latin: face(collection.latin),
      eastAsian: face(collection.eastAsian),
      complexScript: face(collection.complexScript),
    };
    return Object.values(out).some(Boolean) ? out : undefined;
  };
  const major = faces(scheme.majorFont);
  const minor = faces(scheme.minorFont);
  return major || minor ? { major, minor } : undefined;
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
  pres: PresentationOptions,
  slide: SlideOptions,
  slideNumber: number,
  slideCount: number | undefined,
  masters: PresentationOptions["masters"],
  size: { widthPx: number; heightPx: number },
): ProjectedSlide {
  const context = tableContextOf(pres, slide.master, masters);
  const background = slideBackgroundOf(slide, masters, context, size);
  const inheritedChildren = inheritedSlideChildrenOf(slide, masters);
  const children = inheritPlaceholderGeometry(slide.children ?? [], slide, masters, context);
  const textFieldContext = {
    slideNumber,
    slideCount,
    now: new Date(),
    metrics: typeof document === "undefined" ? undefined : browserFontMetrics,
    ...context,
  };
  const inheritedMembers = childMembers(inheritedChildren, IDENTITY, [], textFieldContext).map(
    (member) => {
      delete member.sourceChildIndex;
      return member as ProjectedSlideMember;
    },
  );
  return {
    ...(background ? { background } : {}),
    children,
    members: [...inheritedMembers, ...childMembers(children, IDENTITY, [], textFieldContext)],
  };
}

/** Layout/master shapes drawn beneath a slide. Placeholder shapes are
 * templates and do not render; layout non-placeholders always do, while the
 * master layer respects both the layout's and slide's show-master flags. */
function inheritedSlideChildrenOf(
  slide: SlideOptions,
  masters: PresentationOptions["masters"],
): SlideChild[] {
  const owner = masters?.find((master) => master.name === slide.master) ?? masters?.[0];
  const layout = preferredLayoutOf(slide, owner);
  const children = layout?.children ?? [];
  const masterShapes = slide.showMasterShapes !== false && layout?.showMasterShapes !== false;
  const masterChildren = masterShapes ? (owner?.children ?? []) : [];
  return [...children, ...masterChildren].filter((child) => {
    const shape = "shape" in child ? (child.shape as PlaceholderShape | undefined) : undefined;
    return placeholderTypeOf(shape) === undefined;
  });
}

interface PlaceholderShape {
  placeholder?: string;
  placeholderIndex?: number;
  properties?: ShapePropertiesOptions;
  hidden?: boolean;
  style?: unknown;
  x?: number | string;
  y?: number | string;
  width?: number | string;
  height?: number | string;
  textBody?: TextBodyOptions;
}

/** Resolve a slide shape's missing geometry from its layout placeholder.
 * Office resolves the layout → master placeholder chain for position,
 * shape facets, text defaults and `sz="0"`. */
function inheritPlaceholderGeometry(
  children: SlideChild[],
  slide: SlideOptions,
  masters: PresentationOptions["masters"],
  context: ThemeContext,
): SlideChild[] {
  const owner = masters?.find((master) => master.name === slide.master) ?? masters?.[0];
  const layout = preferredLayoutOf(slide, owner) ?? owner?.layouts?.[0];

  return children.map((child) => {
    const shape = "shape" in child ? (child.shape as PlaceholderShape | undefined) : undefined;
    const placeholderType = placeholderTypeOf(shape);
    if (!shape || !placeholderType) return child;

    const resolved = resolvePlaceholder(placeholderType, layout, owner);
    const next: PlaceholderShape = {
      ...shape,
      ...(resolved.hidden ? { hidden: true } : {}),
    };
    const facets = resolved.facets;
    if (facets) {
      next.properties = {
        ...(facets.geometry ? { geometry: facets.geometry } : {}),
        ...(facets.customGeometry ? { customGeometry: facets.customGeometry } : {}),
        ...(facets.fill ? { fill: facets.fill } : {}),
        ...(facets.outline ? { outline: facets.outline } : {}),
        ...(facets.effects ? { effects: facets.effects } : {}),
        ...(facets.scene3d ? { scene3d: facets.scene3d } : {}),
        ...(facets.shape3d ? { shape3d: facets.shape3d } : {}),
        ...shape.properties,
      };
      if (shape.style === undefined && facets.style) next.style = facets.style;
    }

    if (resolved.position) {
      for (const field of ["x", "y", "width", "height"] as const) {
        if (next[field] === undefined) next[field] = resolved.position[field];
      }
    }
    if (shape.textBody) {
      const inheritedStyles = placeholderTextStylesOf(placeholderType, context);
      const templateBody = facets?.textBody;
      next.textBody = {
        ...shape.textBody,
        bodyProperties: shape.textBody.bodyProperties ?? templateBody?.bodyProperties,
        ...(shape.textBody.anchor == null && templateBody?.anchor != null
          ? { anchor: templateBody.anchor }
          : {}),
        ...(shape.textBody.autoFit == null && templateBody?.autoFit != null
          ? { autoFit: templateBody.autoFit }
          : {}),
        listStyle: mergeTextStyles(
          mergeTextStyles(templateBody?.listStyle, inheritedStyles),
          shape.textBody.listStyle,
        ),
      };
    }

    if (
      !resolved.hidden &&
      (next.x === undefined ||
        next.y === undefined ||
        next.width === undefined ||
        next.height === undefined)
    ) {
      return child;
    }
    return { ...child, shape: next } as SlideChild;
  });
}

function preferredLayoutOf(
  slide: SlideOptions,
  owner: MasterDefinition | undefined,
): LayoutDefinition | undefined {
  const layouts = owner?.layouts ?? [];
  if (slide.layout == null) return undefined;
  return layouts.find(
    (layout) =>
      (layout.layoutId != null && slide.layout === `layout:${layout.layoutId}`) ||
      layout.type === slide.layout ||
      layout.name === slide.layout ||
      layout.matchingName === slide.layout,
  );
}

function placeholderTypeOf(shape: PlaceholderShape | undefined): string | undefined {
  return shape?.placeholder;
}

function placeholderTextStylesOf(
  key: string | undefined,
  context: ThemeContext,
): TextListStyleOptions | undefined {
  if (key === "title" || key === "ctrTitle") return context.textStyles?.title;
  return context.textStyles?.body ?? context.textStyles?.other;
}

function mergeTextStyles(
  explicit: TextListStyleOptions | undefined,
  inherited: TextListStyleOptions | undefined,
): TextListStyleOptions | undefined {
  if (!inherited) return explicit;
  if (!explicit) return inherited;
  const levels = Math.max(explicit.levels?.length ?? 0, inherited.levels?.length ?? 0);
  return {
    ...inherited,
    ...explicit,
    defaultParagraph: explicit.defaultParagraph ?? inherited.defaultParagraph,
    levels: Array.from(
      { length: levels },
      (_, level) =>
        mergeParagraphStyle(
          explicit.levels?.[level] ?? undefined,
          inherited.levels?.[level] ?? undefined,
        ) ?? null,
    ),
  };
}

function mergeParagraphStyle(
  explicit: TextLevelStyle | undefined,
  inherited: TextLevelStyle | undefined,
): TextLevelStyle | undefined {
  if (!inherited) return explicit;
  if (!explicit) return inherited;
  return {
    ...inherited,
    ...explicit,
    defaultRunProperties: {
      ...inherited.defaultRunProperties,
      ...explicit.defaultRunProperties,
    },
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
  const owner = masters?.find((master) => master.name === slide.master) ?? masters?.[0];
  const layout = preferredLayoutOf(slide, owner);
  const background = slide.background ?? layout?.background ?? owner?.background;
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
  color: string | SolidFillOptions | undefined,
  context: ThemeContext,
  phClr: string | undefined,
): ResolvedColor | undefined {
  if (typeof color === "string") {
    const hex = color.replace("#", "").toUpperCase();
    return /^[0-9A-F]{6}$/.test(hex) ? { color: hex, alpha: 1 } : undefined;
  }
  // scRGB/hsl carry channels instead of a token — they resolve to nothing here.
  if (color && "value" in color) {
    const token = color.value;
    if (token === "phClr") return phClr ? transformColor(phClr, transformsOf(color)) : undefined;
    if (/^[0-9A-F]{6}$/i.test(token)) return transformColor(token, transformsOf(color));
    // The raw clrMap tokens spell the first four slots short (bg1/tx1/bg2/tx2);
    // the parsed color map keys them long (background1/text1/…).
    const key = LONG_COLOR_TOKEN[token] ?? token;
    const slot = context.colorMapping[key as keyof ColorMappingOptions] ?? token;
    const hex = context.themeColors?.[slot as keyof ThemeColors];
    return hex ? transformColor(hex, transformsOf(color)) : undefined;
  }
  return undefined;
}

function styleColorOf(
  color: string | SolidFillOptions | undefined,
  context: ThemeContext,
  phClr: string | undefined,
): string | undefined {
  return resolvedStyleColorOf(color, context, phClr)?.color;
}

function stylePaintOf(
  color: string | SolidFillOptions | undefined,
  context: ThemeContext,
  phClr: string | undefined,
): string | undefined {
  const resolved = resolvedStyleColorOf(color, context, phClr);
  if (!resolved) return undefined;
  if (resolved.alpha >= 1) return resolved.color;
  const raw = Number.parseInt(resolved.color, 16);
  return `rgba(${(raw >> 16) & 255}, ${(raw >> 8) & 255}, ${raw & 255}, ${resolved.alpha})`;
}

/** The clrMap's short slot spellings → the parsed color map's long keys. */
const LONG_COLOR_TOKEN: Record<string, string> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
};

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
      color: stylePaintOf(fill.color, context, phClr) ?? fill.color,
    };
  }
  if (fill.type === "gradient") {
    const resolveStops = <T extends { position: number; color: string | SolidFillOptions }>(
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
      foregroundColor: stylePaintOf(fill.foregroundColor, context, phClr) ?? fill.foregroundColor,
      backgroundColor: stylePaintOf(fill.backgroundColor, context, phClr) ?? fill.backgroundColor,
    };
  }
  return fill;
}

/** A SolidFillOptions wrapper or raw EG_ColorChoice → the resolver. */
function transformsOf(
  color: string | SolidFillOptions | undefined,
): ColorTransformOptions | undefined {
  return typeof color === "object" ? color.transforms : undefined;
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
    const stopColor = (color: GradientStopOptions["color"]) => {
      if (typeof color === "string") {
        const paint = color.trim();
        if (/^rgba\(/i.test(paint)) return paint;
        return paint.replace("#", "").toUpperCase();
      }
      if (!color || !("value" in color) || typeof color.value !== "string") return undefined;
      return color.value.replace("#", "").toUpperCase();
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
    if (!src) return undefined;
    if (!fill.tile) return { kind: "image", src };
    return {
      kind: "image",
      src,
      tile: tilePaintOf(fill.tile),
    };
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
