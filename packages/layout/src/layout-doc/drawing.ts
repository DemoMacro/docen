import type { LayoutBlock } from "./block";
import type { LayoutPictureCrop } from "./inline";

// ── floating drawings (anchored shape groups) ──

/** One member's outline stroke (a:ln): width px + hex color plus the line-
 *  dressing tokens (cap/join full-word, dash the prstDash token). */
export interface LayoutDrawingLine {
  px: number;
  color?: string;
  cap?: "round" | "square" | "flat";
  join?: "round" | "bevel" | "miter";
  dash?: string;
}

/** One absolutely-positioned member of a floating drawing. Coordinates and
 *  sizes are px in the drawing's own box (top-left corner the origin) — the
 *  adapter already resolved the group's child coordinate space (chOff/chExt
 *  scaling). */
export type LayoutDrawingMember =
  | {
      kind: "picture";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Renderer media source (data URL); absent → an empty frame. */
      src?: string;
      /** Metafile raster-op emulation blend (SRCPAINT → screen,
       *  SRCAND → multiply) for masked GDI blt pairs. */
      blend?: "screen" | "multiply";
      /** Mirrored content flips (a:xfrm @flipH/@flipV). */
      flipH?: boolean;
      flipV?: boolean;
      /** Source rectangle crop (a:srcRect): the painted region is the
       *  visible remainder of the image edge. */
      crop?: LayoutPictureCrop;
    }
  | {
      kind: "shape";
      x: number;
      y: number;
      width: number;
      height: number;
      /** Preset geometry (a:prstGeom @prst). The renderer maps the presets it
       *  knows and skips the rest; custom geometry stays unprojected. */
      preset?: string;
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Fill opacity 0-1 (the solid color's a:alpha percent ÷ 100);
       *  absent → fully opaque. Fades the fill only, never the stroke. */
      opacity?: number;
      /** Outline stroke: width in px + hex color (absent color → ink). */
      line?: LayoutDrawingLine;
    }
  | {
      kind: "path";
      x: number;
      y: number;
      width: number;
      height: number;
      /** SVG path data in box coordinates (0,0 … width,height) — the adapter
       *  scaled the geometry's own space (custGeom path w/h) into the box. */
      d: string;
      /** Fill rule for self-intersecting outlines (wmf SetPolyFillMode). */
      fillRule?: "evenodd" | "nonzero";
      /** Solid fill, hex RRGGBB; absent → no fill. */
      fill?: string;
      /** Outline stroke (a:ln): width px + hex color + cap/join/dash. */
      line?: LayoutDrawingLine;
    }
  | {
      kind: "textBox";
      x: number;
      y: number;
      width: number;
      height: number;
      /** The shape's own solid fill, hex RRGGBB; absent → no fill (a plain
       *  wps:txbx draws its spPr fill under the text). */
      fill?: string;
      /** Fill opacity 0-1 (the solid color's a:alpha percent ÷ 100). */
      opacity?: number;
      /** The shape's outline stroke (a:ln). Word draws the txbx box even
       *  when the body is empty. */
      line?: LayoutDrawingLine;
      /** Preset geometry (a:prstGeom @prst) — a txbx can live in any shape
       *  (a text-carrying ellipse); the box paints in that shape. */
      preset?: string;
      /** Text insets px (wps:bodyPr lIns/tIns/rIns/bIns, DrawingML defaults
       *  applied by the adapter). */
      insets?: { left?: number; top?: number; right?: number; bottom?: number };
      /** Vertical anchoring (bodyPr @anchor); absent = top (the OOXML
       *  ST_TextAnchoringType default). */
      anchor?: "top" | "center" | "bottom";
      /** bodyPr a:spAutoFit — the shape hugs its text, so vertical anchoring
       *  resolves against the fitted height, not the declared extent (Word
       *  shrinks a stale oversized cy to the text's line height). */
      autoFit?: boolean;
      /** Paint the runs as one unwrapped line at the box origin — metafile
       *  text (GDI ExtTextOut / EMF+ DrawString) draws the string as-is and
       *  never word-wraps, so the width re-fit must not re-break it. */
      nowrap?: boolean;
      /** Clockwise rotation of the text body about the box origin, degrees —
       *  metafile vertical text: a rotated GDI world transform lays runs down
       *  a column; shaping stays horizontal, the paint rotates. */
      rotation?: number;
      /** The box's content as projected blocks — the renderer stacks them
       *  inside the box (the flow never sees them; the drawing wraps none). */
      blocks: LayoutBlock[];
    };

/** A floating shape group (wpg:wgp under a wp:anchor): a drawing box anchored
 *  to its paragraph, members absolutely positioned inside. The flow gives the
 *  anchor paragraph no extra height — the anchor's `relative` axes resolve to
 *  page geometry at paint time (the painter owns the page box).
 *
 *  Word's eight relativeFrom values collapse onto these four semantic axes
 *  per side; the axes the adapter cannot resolve yet (character/line anchor
 *  points, inside/outside under mirrored margins) land on their nearest
 *  axis, a registered gap. */
export interface LayoutDrawingAnchor {
  horizontal: {
    /** column = the content box; leftMargin/rightMargin = its edges; page =
     *  the page box. margin/insideMargin → column, character → column,
     *  outsideMargin → rightMargin (unmirrored). */
    relative: "column" | "leftMargin" | "rightMargin" | "page";
    /** px from the reference box's left edge (posOffset), or a fraction of
     *  the reference box's width (percentOffset / 1000). Exclusive with align. */
    offsetPx?: number;
    percent?: number;
    /** Alignment inside the reference box (inside→left, outside→right). */
    align?: "left" | "center" | "right";
  };
  vertical: {
    /** paragraph = the anchor paragraph's top; topMargin/bottomMargin = the
     *  content box's edges; page = the page box. margin/insideMargin →
     *  topMargin, line → paragraph, outsideMargin → bottomMargin. */
    relative: "paragraph" | "topMargin" | "bottomMargin" | "page";
    offsetPx?: number;
    percent?: number;
    align?: "top" | "center" | "bottom";
  };
}

export interface LayoutDrawing {
  anchor: LayoutDrawingAnchor;
  width: number;
  height: number;
  members: LayoutDrawingMember[];
  /** w:wrap — how text flows around the box. Absent (wrapNone) paints over
   *  or under the text without affecting the flow; "square" shrinks the lines
   *  the box overlaps, "tight" the same lines sliced along `contour` (square's
   *  rectangle without one); "topAndBottom" clears the band. */
  wrap?: "square" | "tight" | "topAndBottom";
  /** ST_WrapSide (w:wrap @side): which side of the box takes text. "right"
   *  moves the wrapped lines past the box's right edge; "left" keeps them on
   *  the left no matter what; "both"/"largest" take the wider side per line
   *  (Word's two-sided wrap approximated — a line packs on one side). */
  wrapSide?: "both" | "left" | "right" | "largest";
  /** wrapTight/through contour (wp:wrapPolygon points) in px, relative to
   *  the drawing box's top-left — the adapter scaled them out of Word's
   *  21600×21600 polygon space. Registered on the zone; the line breaker
   *  slices it per line (a bounding-box approximation without it). */
  contour?: { x: number; y: number }[];
  /** w:behindDoc — painted under the text layer (text strokes stay visible). */
  behind?: boolean;
  /** w:relativeHeight — stacking order within the behind/in-front band
   *  (renderer-only; the layout never reads it). Absent = document order. */
  zIndex?: number;
  /** w:anchor distL/distT/distR/distB in px — how far wrapping text keeps
   *  its distance from the box. square/tight zones widen by left+right,
   *  topAndBottom bands by top+bottom; wrapNone never reads them. */
  distances?: { left?: number; top?: number; right?: number; bottom?: number };
  /** A srcRect-cropped metafile replay clips its members to the extent (GDI
   *  playback semantics — records past the crop line never draw). Absent on
   *  wps text boxes: their text may overflow a stale declared extent. */
  clipMembers?: boolean;
}

/** A drawing's wrap box: its extent grown by the anchor's wrap distances,
 *  clipped to the column. Both zone builders (the flow's post-anchor zones
 *  and the anchor paragraph's own self-zones) share this so the square/
 *  tight/topAndBottom paddings stay in lockstep. `boxTopPx` is the drawing's
 *  own top in the caller's Y space; the distances pad top and bottom around
 *  it. `inCell` applies Word's layoutInCell containment: a cell-anchored
 *  object never extends past its cell — an offset that overflows the cell's
 *  right edge shifts the whole box left to touch it (the wrap zone follows
 *  the shifted box). Body floats keep their raw offset (they may hang into
 *  the margins). Returns undefined when the padded box misses the column
 *  entirely. */
export function drawingWrapBox(
  d: LayoutDrawing,
  boxTopPx: number,
  columnWidth: number,
  inCell = false,
):
  | {
      widthPx: number;
      topPx: number;
      bottomPx: number;
      x0Px: number;
      textAfter?: boolean;
      contour?: { x: number; y: number }[];
    }
  | undefined {
  const left = d.distances?.left ?? 0;
  const offset = inCell
    ? Math.min(Math.max(d.anchor.horizontal.offsetPx ?? 0, 0), Math.max(0, columnWidth - d.width))
    : (d.anchor.horizontal.offsetPx ?? 0);
  const x0 = offset - left;
  const start = Math.max(x0, 0);
  const widthPx =
    Math.min(offset + d.width + left + (d.distances?.right ?? 0), columnWidth) - start;
  if (widthPx <= 0) return undefined;
  // Which side takes the text: "right" forces the right side, "left" keeps
  // the text left no matter what, and both/largest take the wider side —
  // Word's two-sided wrap approximated per line (a line packs on one side).
  const rightSpace = columnWidth - start - widthPx;
  const textAfter =
    d.wrapSide === "right" || (d.wrapSide !== "left" && rightSpace > start) || undefined;
  // The contour rides along in zone coordinates: the drawing box's own
  // top-left sits at (offset-start, distTop) inside the padded zone box.
  const contour = d.contour?.map((p) => ({
    x: p.x + offset - start,
    y: p.y + (d.distances?.top ?? 0),
  }));
  return {
    widthPx,
    topPx: boxTopPx - (d.distances?.top ?? 0),
    bottomPx: boxTopPx + d.height + (d.distances?.bottom ?? 0),
    x0Px: start,
    ...(textAfter ? { textAfter } : {}),
    ...(contour ? { contour } : {}),
  };
}

/** A paragraph's wrapping drawings as flow effects in the caller's Y space
 *  (`baseY` = the anchor paragraph's top; 0 for paragraph-relative zones).
 *  Zones shrink the lines they overlap; bands (topAndBottom, or a square box
 *  covering the full column) clear everything in their band — callers that
 *  cannot dodge a band mid-stack (a paragraph's own lines, a table cell)
 *  drop them. Column/margin/page-anchored and wrapNone drawings produce
 *  neither. */
export function wrapEffectsOf(
  drawings: readonly LayoutDrawing[] | undefined,
  baseY: number,
  columnWidth: number,
  inCell = false,
): { zones: LayoutFloatZone[]; bands: LayoutFloatZone[] } {
  const zones: LayoutFloatZone[] = [];
  const bands: LayoutFloatZone[] = [];
  for (const d of drawings ?? []) {
    if (!d.wrap) continue;
    const { horizontal: h, vertical: v } = d.anchor;
    if (v.relative !== "paragraph" || h.relative !== "column") continue;
    const box = drawingWrapBox(d, baseY + (v.offsetPx ?? 0), columnWidth, inCell);
    if (!box) continue;
    const zone: LayoutFloatZone = {
      widthPx: box.widthPx,
      topPx: box.topPx,
      bottomPx: box.bottomPx,
      x0Px: box.x0Px,
      ...(box.textAfter ? { textAfter: true } : {}),
      ...(box.contour ? { contour: box.contour } : {}),
    };
    if (d.wrap === "topAndBottom" || box.widthPx >= columnWidth - 1) bands.push(zone);
    else zones.push(zone);
  }
  return { zones, bands };
}

// ── blocks ──

/** A floating picture's flow band: text lines overlapping [topPx, bottomPx)
 *  wrap beside it (usable width reduced by widthPx). wrapNone/page/margin
 *  anchors never produce a zone — they sit outside the text flow. */
export interface LayoutFloatZone {
  widthPx: number;
  topPx: number;
  bottomPx: number;
  /** Line-content x of the zone's left edge — pairs with `textAfter` to move
   *  the line's text right of the box (wrapSide right/largest). */
  x0Px?: number;
  /** Text packs RIGHT of the box: the line's usable interval starts past
   *  x0Px + widthPx instead of packing from the left margin. */
  textAfter?: boolean;
  /** Tight/through contour polygon in zone coordinates (the drawing box's
   *  own contour translated by the wrap distances) — the line breaker
   *  slices it at each line's mid-height instead of using the box. */
  contour?: { x: number; y: number }[];
}
