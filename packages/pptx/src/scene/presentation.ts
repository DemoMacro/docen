// The PPTX scene projection: PresentationOptions → flat slide-absolute drawing
// members the core painter paints. Batch 1 covers the vector surface —
// shapes, pictures, lines/connectors, groups, shape text and chart frames —
// through the same format-neutral extraction helpers the docx projection
// shares (@docen/core/geometry). Batch 2 adds graphic-frame tables (a:tbl →
// the frame-table painter's normalized payload). Batch gaps: smartart and
// media-frame children (no member kind / painter wiring yet), a shape's text
// stays unspun when the shape itself rotates, group-level rotation, connector
// presets whose endpoints run reversed, live field evaluation, table styles
// (the tblPr region flags/built-in styles stay unresolved) — each lands
// with its follow-up batch.

import { measureEmu, solidFillOf } from "@docen/core/geometry";
import { emuToPx, type LayoutDrawingMember } from "@docen/layout";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

import { IDENTITY } from "./geometry";
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
  return { widthPx, heightPx, slides: (pres.slides ?? []).map(projectSlide) };
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

function projectSlide(slide: SlideOptions): ProjectedSlide {
  return {
    ...(slide.background ? { background: solidFillOf(slide.background.fill) } : {}),
    members: childMembers(slide.children ?? [], IDENTITY, []),
  };
}
