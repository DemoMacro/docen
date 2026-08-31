import {
  stackBlocks,
  TextMeasurer,
  type LaidOutParagraph,
  type LayoutBlockContext,
  type LayoutDrawing,
  type LayoutDrawingMember,
} from "@docen/layout";
import { Box, Ellipse, Group, Path as LeaferPath, Rect, type IGroup } from "leafer-ui";

import { paintBlock } from "../painter";
import type { DrawingHitBox, PaintColumn, PaintContext } from "./context";
import { addBlendedPictureRun, addCroppedImage, addPlainImage } from "./image";
/** OOXML prstDash tokens → dash patterns in px (line-width units, the host's
 *  preset line styles); unlisted tokens render solid. */
const PRSTDASH_PATTERN: Record<string, number[]> = {
  dot: [1, 3],
  // A 1px-on/1px-off antialiased hairline blends to a faint tint — Word's
  // sysDot boxes read as clear 2px dots at hairline widths (user-verified).
  sysDot: [2, 2],
  dash: [4, 3],
  sysDash: [4, 2],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [4, 2, 1, 2],
  dashDotDot: [4, 3, 1, 3, 1, 3],
  sysDashDotDot: [4, 2, 1, 2, 1, 2],
  lgDash: [12, 3],
  lgDashDot: [12, 3, 1, 3],
  lgDashDotDot: [12, 3, 1, 3, 1, 3],
};
function drawingBoxOf(
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): { x: number; y: number } {
  const { flow } = ctx;
  // The reference box each axis resolves against: the content box (column /
  // topMargin), the page box, an edge (leftMargin/rightMargin/bottomMargin),
  // or — vertically — the anchor paragraph's own top (extent 0: offsets and
  // align both hang off the edge itself). The column's left edge is the
  // CALLER's text column — the page's for body paragraphs, the cell's for
  // cell-anchored floats (Word's layoutInCell), matching the wrap zones the
  // layout computes against the same base.
  const hBox =
    drawing.anchor.horizontal.relative === "page"
      ? { left: 0, width: flow.pageWidthPx }
      : drawing.anchor.horizontal.relative === "rightMargin"
        ? { left: flow.contentLeftPx + flow.contentWidthPx, width: 0 }
        : drawing.anchor.horizontal.relative === "leftMargin"
          ? { left: flow.contentLeftPx, width: 0 }
          : { left: x, width: col?.width ?? flow.contentWidthPx };
  const vBox =
    drawing.anchor.vertical.relative === "page"
      ? { top: 0, height: flow.pageHeightPx }
      : drawing.anchor.vertical.relative === "paragraph"
        ? { top: y, height: 0 }
        : drawing.anchor.vertical.relative === "bottomMargin"
          ? { top: flow.contentTopPx + flow.contentHeightPx, height: 0 }
          : { top: flow.contentTopPx, height: flow.contentHeightPx };
  // Axis position: align inside the reference box, else the offset (px, or a
  // fraction of the reference extent) from its leading edge.
  const axisPos = (
    spec: { offsetPx?: number; percent?: number; align?: string },
    base: number,
    extent: number,
    size: number,
  ): number => {
    if (spec.align === "center") return base + (extent - size) / 2;
    if (spec.align === "right" || spec.align === "bottom") return base + extent - size;
    if (spec.align) return base; // left / top
    if (spec.percent != null) return base + spec.percent * extent;
    return base + (spec.offsetPx ?? 0);
  };
  let boxX = axisPos(drawing.anchor.horizontal, hBox.left, hBox.width, drawing.width);
  const boxY = axisPos(drawing.anchor.vertical, vBox.top, vBox.height, drawing.height);
  // Word's layoutInCell: a cell-anchored object never extends past its cell —
  // an offset that overflows the right edge shifts the whole box left to touch
  // it (the wrap zones shifted with it at layout time). Body floats keep
  // their raw position (they may hang into the margins).
  if (col?.inCell && drawing.anchor.horizontal.relative === "column") {
    boxX = Math.min(Math.max(boxX, hBox.left), hBox.left + Math.max(0, hBox.width - drawing.width));
  }
  return { x: boxX, y: boxY };
}

/** Record a drawing's hit box without painting it — the body pass catalogs
 *  behind-doc floats the earlier pass already painted. */
export function recordDrawingHit(
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  boxes: DrawingHitBox[],
  host: DrawingHost,
): void {
  const box = drawingBoxOf(drawing, x, y, ctx);
  boxes.push({
    page: ctx.pageIndex,
    x: box.x,
    y: box.y,
    width: drawing.width,
    height: drawing.height,
    para: host.para,
    index: host.index,
    kind: "drawing",
  });
}

/** One floating drawing: members absolutely positioned in the drawing's box,
 *  itself placed by the anchor spec against the page geometry. A text box
 *  stacks its own paragraphs inside its insets (the same stackBlocks the
 *  header/footer furniture uses). */
/** Which paragraph carries a drawing, and its position among that paragraph's
 *  drawings (run order — how the PM side re-finds the node). */
interface DrawingHost {
  para: LaidOutParagraph;
  index: number;
}

export function paintDrawing(
  tree: IGroup,
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  col: PaintColumn | undefined,
  host: DrawingHost,
): void {
  const box = drawingBoxOf(drawing, x, y, ctx, col);
  const boxX = box.x;
  const boxY = box.y;
  ctx.hitBoxes?.push({
    page: ctx.pageIndex,
    x: boxX,
    y: boxY,
    width: drawing.width,
    height: drawing.height,
    para: host.para,
    index: host.index,
    kind: "drawing",
  });
  if (drawing.clipMembers) {
    // A srcRect-cropped metafile replay reaches past the extent (GDI clips
    // metafile playback to the rect); wps text boxes must NOT clip — their
    // text may legitimately overflow a stale declared extent. Leafer clips
    // children only on a Box (`overflow` is Box data, Group ignores it).
    const holder = new Box({
      x: boxX,
      y: boxY,
      width: drawing.width,
      height: drawing.height,
      overflow: "hide",
    });
    paintMembers(holder, drawing.members, 0, 0, ctx);
    tree.add(holder);
  } else {
    paintMembers(tree, drawing.members, boxX, boxY, ctx);
  }
}

/** The members of a drawing box (or of an inline picture's metafile replay),
 *  each positioned at its own offset inside the box origin. Shared by the
 *  anchored-drawing and the inline-picture paths — the member shapes are the
 *  same; only the box origin differs. */
export function paintMembers(
  tree: IGroup,
  members: readonly LayoutDrawingMember[],
  boxX: number,
  boxY: number,
  ctx: PaintContext,
): void {
  // A drawing box is a complete little scene: its members paint in full
  // whichever pass anchors the box. A behind-doc watermark's text still lays
  // out as ordinary story rows — threading the behind layer into the member
  // paragraphs would hit paintParagraph's behind branch, which skips line
  // painting entirely, and the body pass never repainted the box.
  const mctx: PaintContext = ctx.layer === "behind" ? { ...ctx, layer: "body" } : ctx;
  for (let i = 0; i < members.length; i++) {
    const m = members[i];
    const mx = boxX + m.x;
    const my = boxY + m.y;
    if (m.kind === "picture" && m.src && !m.crop) {
      // A masked GDI blt sequence (SRCPAINT then SRCAND halves) composites
      // against the metafile's own backdrop. The run is flattened into one
      // image before painting: canvas blend modes only see the editor's
      // layered App canvases, whose destinations are transparent — not the
      // underlying members a ternary raster-op needs.
      const run: Extract<LayoutDrawingMember, { kind: "picture" }>[] = [m];
      let end = i + 1;
      while (end < members.length) {
        const cur = members[end];
        if (cur.kind !== "picture" || !cur.src || cur.crop) break;
        run.push(cur);
        end++;
      }
      if (run.some((p) => p.blend)) {
        // A masked layer is meaningful only inside a composited run: painted
        // alone its opaque mask background (SRCPAINT halves are black-backed)
        // would lay a black slab over the page. A run of one blends against
        // nothing — honest absence beats a wrong slab.
        if (run.length > 1) addBlendedPictureRun(tree, run, boxX, boxY, ctx);
        i = end - 1;
        continue;
      }
    }
    if (m.kind === "picture") {
      if (m.src && m.crop) {
        addCroppedImage(tree, m.src, m.crop, mx, my, m.width, m.height, ctx, m.flipH, m.flipV);
      } else if (m.src) {
        addPlainImage(tree, m, mx, my, ctx);
      } else {
        tree.add(
          new Rect({
            x: mx,
            y: my,
            width: m.width,
            height: m.height,
            fill: "#f3f3f3",
            stroke: "#c4c4c4",
            strokeWidth: 1,
          }),
        );
      }
    } else if (m.kind === "path") {
      tree.add(
        new LeaferPath({
          x: mx,
          y: my,
          width: m.width,
          height: m.height,
          // Leafer's Path takes SVG path data under `path` (its `data` holds
          // the parsed command array — a string there paints nothing).
          path: m.d,
          fill: m.fill ? `#${m.fill}` : undefined,
          stroke: m.line ? (m.line.color ? `#${m.line.color}` : "#000000") : undefined,
          // A dashed hairline disappears entirely: the antialiased 1 px stroke
          // fades under the dash gaps and nothing survives. Hold a 1.5 px
          // floor on dashed strokes so the pattern renders; solid lines keep
          // their true width (fading is invisible there).
          strokeWidth:
            m.line?.px != null ? (m.line.dash ? Math.max(m.line.px, 1.5) : m.line.px) : undefined,
          strokeCap:
            m.line?.cap === "round" ? "round" : m.line?.cap === "square" ? "square" : undefined,
          strokeJoin:
            m.line?.join === "round" ? "round" : m.line?.join === "bevel" ? "bevel" : undefined,
          dashPattern: m.line?.dash ? PRSTDASH_PATTERN[m.line.dash] : undefined,
          // Leafer spells the SVG fill-rule attribute `windingRule`.
          windingRule: m.fillRule,
        }),
      );
    } else if (m.kind === "shape") {
      paintShapeBox(tree, { ...m, x: mx, y: my }, false);
    } else {
      // The txbx shape's own paint (prstGeom silhouette + spPr fill + a:ln)
      // sits under its text — Word draws the box even when the body is empty
      // (a plain text box is white fill + an accent hairline, visible on a
      // white page).
      paintShapeBox(tree, { ...m, x: mx, y: my }, true);
      const measurer = new TextMeasurer(ctx.metrics);
      const left = m.insets?.left ?? 0;
      // Metafile runs carry nowrap: GDI draws the string as-is, so the width
      // re-fit must not re-break it into phantom lines.
      const inner = m.nowrap
        ? Number.POSITIVE_INFINITY
        : Math.max(0, m.width - left - (m.insets?.right ?? 0));
      // A text box shares its STORY's doc grid with the surrounding text: a
      // body box snaps to the section grid and centers its grid rows
      // (onGrid — the half-leading the reference renders), while a
      // header/footer box gets no pitch at all (the furniture paint context
      // clears it — the story keeps natural line heights). bodyPr
      // @compatLnSpc plays no role: Word ignores it for wps txbxContent.
      // Metafile text carries nowrap: GDI strings are absolutely positioned
      // by the replay (baseline-derived y), not story rows — the grid pad
      // would shove them half a pitch off their drawn spot.
      const grid: LayoutBlockContext | undefined =
        ctx.flow.linePitchPx && !m.nowrap
          ? { linePitchPx: ctx.flow.linePitchPx, onGrid: true }
          : undefined;
      const laid = stackBlocks(m.blocks, inner, grid, measurer);
      let oy = m.insets?.top ?? 0;
      if (m.anchor === "center" || m.anchor === "bottom") {
        // spAutoFit shrinks the drawn box to the text, so slack resolves
        // against the fitted height — an oversized declared extent (stale
        // cy from a template) must not push the text down/center the box.
        const boxH = m.autoFit
          ? (m.insets?.top ?? 0) + laid.heightPx + (m.insets?.bottom ?? 0)
          : m.height;
        const slack = boxH - (m.insets?.top ?? 0) - (m.insets?.bottom ?? 0) - laid.heightPx;
        oy += m.anchor === "center" ? Math.max(0, slack / 2) : Math.max(0, slack);
      }
      if (m.rotation) {
        // Vertical metafile text (a rotated GDI world transform): the body
        // shapes horizontally as usual, then rotates about the box origin —
        // a group keeps the offsets in text space and carries the angle. A
        // clockwise 90° (vertical punctuation) swings the box left of the
        // origin, so the group shifts right by the box width and the ink
        // anchors the cell's top-right corner. Any other angle (Word's
        // diagonal watermark) pivots about the box CENTER — the anchor box
        // was laid centered on the page, so the rotated box stays centered.
        const pivot = m.rotation === 90;
        const group = new Group({
          x: pivot ? mx + m.width : mx + m.width / 2,
          y: pivot ? my : my + m.height / 2,
          rotation: m.rotation,
        });
        const dx = pivot ? 0 : -m.width / 2;
        const dy = pivot ? 0 : -m.height / 2;
        for (const item of laid.stack) {
          paintBlock(group, item.block, left + dx, oy + dy + item.yPx, mctx, {
            width: inner,
            inCell: true,
          });
        }
        tree.add(group);
      } else {
        for (const item of laid.stack) {
          paintBlock(tree, item.block, mx + left, my + oy + item.yPx, mctx, {
            width: inner,
            inCell: true,
          });
        }
      }
    }
  }
}

/** Hex RRGGBB + opacity 0-1 → a CSS rgba color (Leafer parses CSS strings;
 *  the alpha must ride on the fill, not element opacity, so a stroke on the
 *  same shape stays opaque). */
function rgbaOf(hex: string, opacity: number): string {
  const n = parseInt(hex, 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${opacity})`;
}

/** One box-like shape's own paint — preset silhouette + solid fill + outline
 *  stroke. Shared by standalone shape members and text-box shapes (a txbx is
 *  a shape carrying text; Word paints its prstGeom under the body). Unknown
 *  presets degrade to a plain rectangle when `rectFallback` (a txbx is a box
 *  by nature) and skip otherwise (an honest absence). */
function paintShapeBox(
  tree: IGroup,
  box: {
    x: number;
    y: number;
    width: number;
    height: number;
    preset?: string;
    fill?: string;
    opacity?: number;
    line?: { px: number; color?: string; dash?: string };
  },
  rectFallback: boolean,
): void {
  if (
    box.preset != null &&
    box.preset !== "rect" &&
    box.preset !== "roundRect" &&
    box.preset !== "ellipse"
  ) {
    if (!rectFallback) return;
  }
  const fill = box.fill
    ? box.opacity != null
      ? rgbaOf(box.fill, box.opacity)
      : `#${box.fill}`
    : undefined;
  const stroke = box.line ? (box.line.color ? `#${box.line.color}` : "#000000") : undefined;
  // A dashed hairline vanishes in the dash gaps — hold the same 1.5 px floor
  // the member path uses so preset dashes render.
  const strokeWidth =
    box.line?.px != null && box.line.dash ? Math.max(box.line.px, 1.5) : box.line?.px;
  const common = {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    fill,
    stroke,
    strokeWidth,
    // Closed shapes default to an inside stroke, under which Leafer's dash
    // pass paints nothing — center stroke renders the dashPattern.
    strokeAlign: "center",
    dashPattern: box.line?.dash ? PRSTDASH_PATTERN[box.line.dash] : undefined,
  };
  if (box.preset === "ellipse") {
    tree.add(new Ellipse(common));
    return;
  }
  tree.add(
    new Rect({
      ...common,
      cornerRadius: box.preset === "roundRect" ? Math.min(box.width, box.height) / 6 : undefined,
    }),
  );
}
