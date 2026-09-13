// Slide painting shared by the main strip and the thumbnails panel, plus the
// panel's constants. The deck paints once per surface: both call sites hand
// in their own Leafer tree and let this loop place one background + members
// group per slide; the thumbnails shrink through the tree's own scale.

import { paintMembers } from "@docen/core";
import { browserFontMetrics } from "@docen/layout";
import type { ProjectedPresentation } from "@docen/pptx";
import { Group, Rect, type IGroup } from "leafer-ui";

/** Gap between consecutive slides in the main strip, px. */
export const SLIDE_GAP_PX = 24;
/** Thumbnail surface width, px — the panel's fixed column width. */
export const THUMB_WIDTH_PX = 160;
/** Visual gap between thumbnails, px (screen space, not slide space). */
export const THUMB_GAP_PX = 12;

/** Paint every slide into `tree` top-down. `pitch`/`yStart` are in slide
 *  coordinates: the main strip uses the slide height + its px gap, the
 *  thumbnails pass their gap divided by the tree scale so the visual
 *  spacing lands right after scaling. */
export function paintSlideDeck(
  tree: IGroup,
  pres: ProjectedPresentation,
  rerender: () => void,
  yPitch = pres.heightPx + SLIDE_GAP_PX,
  yStart = SLIDE_GAP_PX,
): void {
  let y = yStart;
  for (let i = 0; i < pres.slides.length; i++) {
    const slide = pres.slides[i]!;
    const slideGroup = new Group({ x: 0, y });
    slideGroup.add(
      new Rect({
        width: pres.widthPx,
        height: pres.heightPx,
        fill: slide.background ? `#${slide.background}` : "#ffffff",
        stroke: "#c4c4c4",
        strokeWidth: 1,
      }),
    );
    paintMembers(slideGroup, slide.members, 0, 0, {
      metrics: browserFontMetrics,
      flow: {
        pageWidthPx: pres.widthPx,
        pageHeightPx: pres.heightPx,
        contentWidthPx: pres.widthPx,
        contentHeightPx: pres.heightPx,
        contentLeftPx: 0,
        contentTopPx: 0,
      },
      pageIndex: i,
      pageCount: pres.slides.length,
      layer: "body",
      rerender,
    });
    tree.add(slideGroup);
    y += yPitch;
  }
}
