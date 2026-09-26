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
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

import { IDENTITY } from "./geometry";
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
  /** The slide's solid background fill, hex RRGGBB; absent → the painter's
   *  page default. Gradient/picture backgrounds stay unprojected (batch gap). */
  background?: string;
  /** Slide-absolute drawing members, paint order = document order. */
  members: LayoutDrawingMember[];
}

/** Project a parsed presentation into the paintable shape: sizes are resolved
 *  to px and every slide's children flatten into members. Pure — the input is
 *  not mutated, so callers may re-project on demand. */
export function projectPresentation(pres: PresentationOptions): ProjectedPresentation {
  const { widthPx, heightPx } = slideSizePx(pres.size);
  const context = tableContextOf(pres);
  return {
    widthPx,
    heightPx,
    slides: (pres.slides ?? []).map((slide, index) =>
      projectSlide(slide, index + 1, pres.slides?.length, context),
    ),
  };
}

/** The theme/table-style context every slide's projection shares: the first
 * master's color scheme plus the tblStyleLst entries keyed by GUID. */
function tableContextOf(pres: PresentationOptions): {
  themeColors?: ThemeColors;
  tableStyles?: Record<string, ReturnType<typeof regionsOf>>;
} {
  const scheme = pres.masters?.[0]?.theme?.colorScheme;
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
  return { themeColors, tableStyles };
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
  context: {
    themeColors?: ThemeColors;
    tableStyles?: Record<string, ReturnType<typeof regionsOf>>;
  },
): ProjectedSlide {
  return {
    ...(slide.background ? { background: solidFillOf(slide.background.fill) } : {}),
    members: childMembers(slide.children ?? [], IDENTITY, [], {
      slideNumber,
      slideCount,
      now: new Date(),
      ...context,
    }),
  };
}
