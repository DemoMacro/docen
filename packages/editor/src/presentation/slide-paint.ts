// Slide painting shared by the main strip and the thumbnails panel, plus the
// panel's constants. The deck paints once per surface: both call sites hand
// in their own Leafer tree and let this loop place one background + members
// group per slide; the thumbnails shrink through the tree's own scale.

import { paintMembers } from "@docen/core";
import { browserFontMetrics } from "@docen/layout";
import type {
  ProjectedSlideMember,
  ProjectedPresentation,
  ProjectedSlide,
  ProjectedSlideBackground,
} from "@docen/pptx";
import { Group, Rect, type IGroup } from "leafer-ui";

import { memberBatches, paintSignature } from "./slide-diff";

/** Gap between consecutive slides in the main strip, px. */
export const SLIDE_GAP_PX = 24;

/** Thumbnail surface width, px — the panel's fixed column width. */
export const THUMB_WIDTH_PX = 160;
/** Visual gap between thumbnails, px (screen space, not slide space). */
export const THUMB_GAP_PX = 12;

/** A projected background color → CSS: canonical hex gains #; an alpha
 *  transform has already resolved to rgba(). */
function cssColorOf(color: string): string {
  return /^rgba\(/i.test(color) ? color : `#${color}`;
}

function paintMemberBatch(
  slideGroup: IGroup,
  batch: readonly ProjectedSlideMember[],
  pres: ProjectedPresentation,
  rerender: () => void,
  index: number,
  at?: number,
): void {
  const group = new Group();
  slideGroup.addAt(group, at ?? slideGroup.children.length);
  paintMembers(group, batch, 0, 0, {
    metrics: browserFontMetrics,
    flow: {
      pageWidthPx: pres.widthPx,
      pageHeightPx: pres.heightPx,
      contentWidthPx: pres.widthPx,
      contentHeightPx: pres.heightPx,
      contentLeftPx: 0,
      contentTopPx: 0,
    },
    pageIndex: index,
    pageCount: pres.slides.length,
    layer: "body",
    rerender,
  });
}

/** The slide's background → the Leafer fill: solid hex, a linear/radial
 *  gradient (OOXML's angle is clockwise from east, screen y-down — the same
 *  sweep), or the picture fill stretched over the slide (PowerPoint's
 *  default picture background behavior). */
function slideFillOf(
  bg: ProjectedSlideBackground | undefined,
  widthPx: number,
  heightPx: number,
): string | Record<string, unknown> {
  if (!bg) return "#ffffff";
  if (bg.kind === "solid") return cssColorOf(bg.color);
  if (bg.kind === "image") {
    if (!bg.tile) return { type: "image", url: bg.src, mode: "stretch" };
    return {
      type: "image",
      url: bg.src,
      ...bg.tile,
    };
  }
  const stops = bg.stops.map((stop) => ({
    offset: stop.position / 100,
    color: cssColorOf(stop.color),
  }));
  if (bg.path) {
    // Radial: the focus sits center, the rim reaches the box edge (the
    // "to" point sets the radius).
    return {
      type: "radial",
      from: { x: widthPx / 2, y: heightPx / 2 },
      to: { x: widthPx / 2, y: heightPx },
      stops,
    };
  }
  const theta = ((bg.angle ?? 0) * Math.PI) / 180;
  // The gradient line spans the full box: its length is the box extents
  // projected on the direction.
  const length = widthPx * Math.abs(Math.cos(theta)) + heightPx * Math.abs(Math.sin(theta));
  return {
    type: "linear",
    from: {
      x: widthPx / 2 - (Math.cos(theta) * length) / 2,
      y: heightPx / 2 - (Math.sin(theta) * length) / 2,
    },
    to: {
      x: widthPx / 2 + (Math.cos(theta) * length) / 2,
      y: heightPx / 2 + (Math.sin(theta) * length) / 2,
    },
    stops,
  };
}

/** Paint one slide (background + members) as a group at strip position `y`.
 *  `index` feeds the paint context's page bookkeeping. */
function paintSlideGroup(
  pres: ProjectedPresentation,
  slide: ProjectedPresentation["slides"][number],
  y: number,
  rerender: () => void,
  index: number,
): Group {
  const slideGroup = new Group({ x: 0, y });
  slideGroup.add(
    new Rect({
      width: pres.widthPx,
      height: pres.heightPx,
      fill: slideFillOf(slide.background, pres.widthPx, pres.heightPx),
      stroke: "#c4c4c4",
      strokeWidth: 1,
    }),
  );
  for (const batch of memberBatches(slide.members)) {
    paintMemberBatch(slideGroup, batch, pres, rerender, index);
  }
  return slideGroup;
}

/** Replace one slide while reusing Leafer nodes whose projected paint is
 *  unchanged. Children are grouped by source-child index, so moving one
 *  textbox does not repaint unrelated pictures, tables, charts or backgrounds. */
export function repaintSlide(
  tree: IGroup,
  pres: ProjectedPresentation,
  index: number,
  rerender: () => void,
  y: number,
  previous?: ProjectedSlide,
): void {
  const slide = pres.slides[index]!;
  const oldGroup = tree.children[index] as IGroup | undefined;
  const oldBatches = previous ? memberBatches(previous.members) : [];
  const newBatches = memberBatches(slide.members);
  const backgroundChanged =
    !oldGroup ||
    !previous ||
    oldBatches.length !== newBatches.length ||
    oldGroup.children.length !== oldBatches.length + 1 ||
    paintSignature(previous.background) !== paintSignature(slide.background);
  if (backgroundChanged) {
    const oldIndex = oldGroup ? tree.children.indexOf(oldGroup) : -1;
    oldGroup?.remove();
    const group = paintSlideGroup(pres, slide, y, rerender, index);
    if (oldIndex < 0) tree.add(group);
    else tree.addAt(group, oldIndex);
    return;
  }
  oldGroup!.y = y;
  newBatches.forEach((batch, batchIndex) => {
    const old = oldGroup!.children[batchIndex + 1] as IGroup | undefined;
    const unchanged = old && paintSignature(batch) === paintSignature(oldBatches[batchIndex] ?? []);
    if (unchanged) return;
    old?.remove();
    paintMemberBatch(oldGroup!, batch, pres, rerender, index, batchIndex + 1);
  });
}

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
  for (let i = 0; i < pres.slides.length; i++) {
    tree.add(paintSlideGroup(pres, pres.slides[i]!, yStart + i * yPitch, rerender, i));
  }
}
