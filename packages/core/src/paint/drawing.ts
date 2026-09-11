import {
  anchorAxisPos,
  stackBlocks,
  TextMeasurer,
  type LaidOutParagraph,
  type LayoutBlockContext,
  type LayoutDrawing,
  type LayoutDrawingLine,
  type LayoutDrawingMember,
  type LayoutDrawingShadow,
} from "@docen/layout";
import { Box, Ellipse, Group, Path as LeaferPath, Rect, type IGroup } from "leafer-ui";

import { paintBlock } from "../painter";
import { paintChartMember } from "./chart";
import type { DrawingHitBox, PaintColumn, PaintContext } from "./context";
import { addBlendedPictureRun, addCroppedImage, addPlainImage, shadowEffectOf } from "./image";
import { strokePropsOf } from "./line";
function drawingBoxOf(
  drawing: LayoutDrawing,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): { x: number; y: number } {
  const { flow } = ctx;
  // The flow pinned this box at the drawing's first resolution (the two-pass
  // wrap): the page replay can move the anchor paragraph, but the box the
  // text wrapped around stays put — paint there, not at the moved anchor.
  if (drawing.pinned) return { ...drawing.pinned };
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
  // Axis position — the shared resolution (the flow's wrap zones compute the
  // same bases, so painter placement and text avoidance cannot drift).
  let boxX = anchorAxisPos(drawing.anchor.horizontal, hBox.left, hBox.width, drawing.width);
  const boxY = anchorAxisPos(drawing.anchor.vertical, vBox.top, vBox.height, drawing.height);
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
  const origin = ctx.origin;
  const box = drawingBoxOf(drawing, x + (origin?.x ?? 0), y + (origin?.y ?? 0), ctx);
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
 *  drawings (run order — how the PM side re-finds the node). `kind` says
 *  which hit sequence the position counts ("drawing" = floating, "inline" =
 *  the paragraph's inline pictures/charts — an inline chart member's
 *  sub-elements register under it too). */
interface DrawingHost {
  para: LaidOutParagraph;
  index: number;
  kind: "drawing" | "inline";
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
  // The anchor resolves against page-local positions (the flow's absolute
  // bases and the caller's anchor-paragraph position), while the members
  // paint group-local below — ctx.origin is the gap between the two.
  const offX = ctx.origin?.x ?? 0;
  const offY = ctx.origin?.y ?? 0;
  const box = drawingBoxOf(drawing, x + offX, y + offY, ctx, col);
  const boxX = box.x - offX;
  const boxY = box.y - offY;
  ctx.hitBoxes?.push({
    page: ctx.pageIndex,
    x: box.x,
    y: box.y,
    width: drawing.width,
    height: drawing.height,
    para: host.para,
    index: host.index,
    kind: "drawing",
    ...(drawing.rotation ? { rotation: drawing.rotation } : {}),
    ...(ctx.layer === "behind" ? { behind: true } : {}),
  });
  // The members' target and origin: the page tree at the box origin — or, on
  // a rotated drawing, a group parked at the box CENTER carrying the angle
  // (same pivot trick as the rotated text-box member), with the content
  // re-offset so the box stays centered under the rotation.
  let target: IGroup = tree;
  let ox = boxX;
  let oy = boxY;
  if (drawing.rotation) {
    const spinner = new Group({
      x: boxX + drawing.width / 2,
      y: boxY + drawing.height / 2,
      rotation: drawing.rotation,
    });
    tree.add(spinner);
    target = spinner;
    ox = -drawing.width / 2;
    oy = -drawing.height / 2;
  }
  if (drawing.flipH || drawing.flipV) {
    // A mirrored drawing rides a negative-scale group inside (any) spinner:
    // the origin shifts to the far edge so the mirrored content lands back
    // inside the same extent box (the plainImageLeaf flip trick), and the
    // members re-anchor to the group's own origin.
    const mirror = new Group({
      x: drawing.flipH ? ox + drawing.width : ox,
      y: drawing.flipV ? oy + drawing.height : oy,
      ...(drawing.flipH ? { scaleX: -1 } : {}),
      ...(drawing.flipV ? { scaleY: -1 } : {}),
    });
    target.add(mirror);
    target = mirror;
    ox = 0;
    oy = 0;
  }
  if (drawing.clipMembers) {
    // A srcRect-cropped metafile replay reaches past the extent (GDI clips
    // metafile playback to the rect); wps text boxes must NOT clip — their
    // text may legitimately overflow a stale declared extent. Leafer clips
    // children only on a Box (`overflow` is Box data, Group ignores it).
    const holder = new Box({
      x: ox,
      y: oy,
      width: drawing.width,
      height: drawing.height,
      overflow: "hide",
    });
    // The holder's page origin rides along: chart boxes register in page
    // space while the members paint tree-local inside the clip box.
    paintMembers(holder, drawing.members, 0, 0, ctx, host, ox + offX, oy + offY);
    target.add(holder);
  } else {
    // A rotated or mirrored drawing transforms in a group — its text lines'
    // screen geometry no longer matches the stack, so the caret map gets no
    // stack to register.
    paintMembers(
      target,
      drawing.members,
      ox,
      oy,
      ctx,
      drawing.rotation || drawing.flipH || drawing.flipV ? undefined : host,
      // Members paint directly at mx = boxX + m.x — their position already
      // carries the box origin, so the page-space top-up is just ctx.origin.
      offX,
      offY,
    );
  }
}

/** The members of a drawing box (or of an inline picture's metafile replay),
 *  each positioned at its own offset inside the box origin. Shared by the
 *  anchored-drawing and the inline-picture paths — the member shapes are the
 *  same; only the box origin differs. `host` (an anchored drawing only) lets
 *  an editable text-box member register its laid stack with the caret map.
 *  `originX/Y` is the page-space offset the member's own position does NOT
 *  carry: holder-painted members (clip box, inline replay) position
 *  tree-local, so callers pass the holder's page origin; members painted
 *  directly at mx = boxX + m.x already carry it and pass ctx.origin alone.
 *  Chart hit boxes register in page space and use it verbatim. */
export function paintMembers(
  tree: IGroup,
  members: readonly LayoutDrawingMember[],
  boxX: number,
  boxY: number,
  ctx: PaintContext,
  host?: DrawingHost,
  originX = 0,
  originY = 0,
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
    // A group member registers its own hit box (painted after the group's, so
    // the member wins the click — whether that selects the member or falls
    // back to the group is the editor's state call). `host` gates it: inline
    // picture replays carry no host, and a rotated drawing's members paint in
    // spinner space where their box geometry no longer matches the page.
    if (host && m.childPath) {
      ctx.hitBoxes?.push({
        page: ctx.pageIndex,
        x: mx + (ctx.origin?.x ?? 0),
        y: my + (ctx.origin?.y ?? 0),
        width: m.width,
        height: m.height,
        para: host.para,
        index: host.index,
        kind: host.kind,
        childPath: m.childPath,
        ...(ctx.layer === "behind" ? { behind: true } : {}),
      });
    }
    if (m.kind === "picture" && m.src && !m.crop) {
      // A masked GDI blt sequence (SRCPAINT then SRCAND halves) composites
      // against the metafile's own backdrop. The run is flattened into one
      // image before painting: canvas blend modes only see the editor's
      // layered App canvases, whose destinations are transparent — not the
      // underlying members a ternary raster-op needs.
      const run: Extract<LayoutDrawingMember, { kind: "picture" }>[] = [m];
      let runHasBlend = !!m.blend;
      let end = i + 1;
      while (end < members.length) {
        const cur = members[end];
        if (cur.kind !== "picture" || !cur.src || cur.crop) break;
        run.push(cur);
        runHasBlend ||= !!cur.blend;
        end++;
      }
      if (runHasBlend) {
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
        addCroppedImage(tree, m.src, m.crop, mx, my, m.width, m.height, ctx, m.flipH, m.flipV, {
          filter: m.filter,
          opacity: m.opacity,
          shadow: m.shadow,
          line: m.line,
        });
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
            ...(m.line ? strokePropsOf(m.line) : { stroke: "#c4c4c4", strokeWidth: 1 }),
            strokeAlign: "center",
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
          ...strokePropsOf(m.line),
          // Adjacent same-color fills share their edge and the rasterizer
          // leaves a 1px antialiasing seam between them — a hairline in the
          // fill color closes it (an outlined member keeps its own stroke).
          ...(m.fill && !m.line ? { stroke: `#${m.fill}`, strokeWidth: 1 } : {}),
          // Leafer spells the SVG fill-rule attribute `windingRule`.
          windingRule: m.fillRule,
          ...shadowEffectOf(m.shadow),
        }),
      );
    } else if (m.kind === "chart") {
      // The sub-element boxes ride the host gate (a rotated drawing's members
      // paint in spinner space — no page-space geometry to register) and skip
      // group interiors: a chart inside a group selects as the member through
      // its own childPath box, Word's group granularity.
      paintChartMember(
        tree,
        { ...m, x: mx, y: my },
        host && !m.childPath
          ? {
              ctx,
              para: host.para,
              index: host.index,
              kind: host.kind,
              // originX/Y already is the holder's page origin (the callers
              // fold ctx.origin in) — adding it again would shift every
              // sub-element box by the content inset.
              ox: originX,
              oy: originY,
            }
          : undefined,
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
      // re-fit must not re-break it into phantom lines. A vertical body
      // (bodyPr @vert) lays out against the transposed column — the box
      // height less the top/bottom insets (see the textVertical branch).
      const inner = m.nowrap
        ? Number.POSITIVE_INFINITY
        : m.textVertical
          ? Math.max(0, m.height - (m.insets?.top ?? 0) - (m.insets?.bottom ?? 0))
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
      if (!m.textVertical && (m.anchor === "center" || m.anchor === "bottom")) {
        // spAutoFit shrinks the drawn box to the text, so slack resolves
        // against the fitted height — an oversized declared extent (stale
        // cy from a template) must not push the text down/center the box.
        const boxH = m.autoFit
          ? (m.insets?.top ?? 0) + laid.heightPx + (m.insets?.bottom ?? 0)
          : m.height;
        const slack = boxH - (m.insets?.top ?? 0) - (m.insets?.bottom ?? 0) - laid.heightPx;
        oy += m.anchor === "center" ? Math.max(0, slack / 2) : Math.max(0, slack);
      }
      // An editable body registers its laid stack with the caret map (a double
      // click edits the text in place). Excluded: metafile text (drawn GDI
      // art — nowrap) and rotated/vertical stacks (group-space geometry — the
      // caret map has no transposed position model yet). A group interior
      // registers too — the host's childPath re-finds the member's wpsShape
      // node.
      if (host && !m.nowrap && !m.rotation && !m.textVertical && ctx.shapeTextStacks) {
        ctx.shapeTextStacks.push({
          page: ctx.pageIndex,
          host: { para: host.para, index: host.index, childPath: m.childPath },
          xPx: mx + left + (ctx.origin?.x ?? 0),
          yPx: my + oy + (ctx.origin?.y ?? 0),
          items: laid.stack,
        });
      }
      // Metafile text is drawn art replayed from GDI records, not editable
      // story rows: its strings carry no paragraph marks, so the marks pass
      // must not paint a ¶ into the box (a wps txbx keeps them — Word shows
      // marks inside text boxes too).
      const tctx = m.nowrap ? { ...mctx, showMarks: false } : mctx;
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
          paintBlock(group, item.block, left + dx, oy + dy + item.yPx, tctx, {
            width: inner,
            inCell: true,
            shapeText: true,
          });
        }
        tree.add(group);
      } else if (m.textVertical) {
        // bodyPr @vert: the body shaped horizontally (against the transposed
        // column = the box height) rotates into place. "vertical" (Word's
        // Rotate all text 90°) reads top-down with columns advancing
        // right-to-left — the group anchors the first column's top-right text
        // origin and the clockwise 90° swing lays each laid line down and each
        // next column left. "vertical270" reads bottom-up with columns
        // left-to-right — anchored at the bottom-left origin, swung counter-
        // clockwise. `oy` above was skipped, as the rotation branch skips it.
        const vert = m.textVertical === "vertical";
        const rIns = m.insets?.right ?? 0;
        const bIns = m.insets?.bottom ?? 0;
        // Vertical anchoring resolves along the column-stack direction (the
        // box width's slack over the laid stack height).
        const slack = Math.max(0, m.width - left - rIns - laid.heightPx);
        const lead = m.anchor === "center" ? slack / 2 : m.anchor === "bottom" ? slack : 0;
        const group = new Group({
          x: vert ? mx + m.width - rIns - lead : mx + left + lead,
          y: vert ? my + (m.insets?.top ?? 0) : my + m.height - bIns,
          rotation: vert ? 90 : -90,
        });
        for (const item of laid.stack) {
          paintBlock(group, item.block, 0, item.yPx, tctx, {
            width: inner,
            inCell: true,
            shapeText: true,
          });
        }
        tree.add(group);
      } else {
        for (const item of laid.stack) {
          paintBlock(tree, item.block, mx + left, my + oy + item.yPx, tctx, {
            width: inner,
            inCell: true,
            shapeText: true,
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
    line?: LayoutDrawingLine | { px: number; color?: string; dash?: string };
    shadow?: LayoutDrawingShadow;
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
  const common = {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    fill,
    ...strokePropsOf(box.line),
    ...shadowEffectOf(box.shadow),
    // Closed shapes default to an inside stroke, under which Leafer's dash
    // pass paints nothing — center stroke renders the dashPattern.
    strokeAlign: "center",
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
