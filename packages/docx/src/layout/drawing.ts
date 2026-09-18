// Floating drawings: wpg groups (nested groups flattened through their
// child coordinate space), standalone wps shape runs, and floating pictures
// — each anchored to its paragraph through the shared wp:anchor mapping.

import {
  arcToSegments,
  fillOpacityOf,
  lineEndMembersOf,
  linePathData,
  outlineOf,
  outerShadowOf,
  presetShapePaths,
  presetShapeTextRect,
  solidFillOf,
} from "@docen/core/geometry";

// The picture-border projection (runs.ts) reads the same stroke extractor.
export { outlineOf };
import {
  emuToPx,
  type LayoutDrawing,
  type LayoutDrawingAnchor,
  type LayoutDrawingMember,
  type LayoutDrawingShadow,
  type LayoutParagraph,
} from "@docen/layout";
import type { GeometryGuide } from "@office-open/core";
import type { CustomGeometryOptions } from "@office-open/core/drawing";
import type { GroupChildMediaData, GroupOptions, MediaDataTransformation } from "@office-open/docx";

import type { ProjectContext } from "./context";
import { isRecord, measureEmu, num, str, type BodyParagraph, type Rec } from "./guards";
import { metafileMembers, pictureSrc } from "./media";
import { projectParagraph } from "./paragraph";

// ── floating drawings (wpg group runs) ──

/** DrawingML text-inset defaults (a:bodyPr), EMU. */
const BODY_INSET_EMU = { left: 91440, right: 91440, top: 45720, bottom: 45720 };

/** A picture's adjustment effects → the renderer description: the percent
 *  blip effects (luminance/hsl/grayscale/blur) map into a CSS-filter
 *  composite string; the shape's outer shadow resolves distance+direction
 *  into a px offset. The alpha modulate comes back separately (the leaf
 *  fades natively, no pixels); the color-mapping effects the filter grammar
 *  cannot express (duotone, bi-level, tint, color replace) stay unprojected. */
export function pictureAdjustOf(
  pic: Rec,
): { filter?: string; opacity?: number; shadow?: LayoutDrawingShadow } | undefined {
  const effects = pic.blipEffects;
  if (!isRecord(effects) && !isRecord(pic.effects)) return undefined;
  const css: string[] = [];
  let opacity: number | undefined;
  if (isRecord(effects)) {
    const lum = isRecord(effects.luminance) ? effects.luminance : undefined;
    if (lum) {
      const bright = num(lum.bright);
      const contrast = num(lum.contrast);
      if (bright) css.push(`brightness(${(1 + bright / 100).toFixed(2)})`);
      if (contrast) css.push(`contrast(${(1 + contrast / 100).toFixed(2)})`);
    }
    const hsl = isRecord(effects.hsl) ? effects.hsl : undefined;
    if (hsl) {
      const hue = num(hsl.hue);
      const saturation = num(hsl.saturation);
      if (hue) css.push(`hue-rotate(${hue}deg)`);
      if (saturation) css.push(`saturate(${(1 + saturation / 100).toFixed(2)})`);
    }
    if (effects.grayscale === true) css.push("grayscale(1)");
    const blur = isRecord(effects.blur) ? num(effects.blur.radius) : undefined;
    if (blur) css.push(`blur(${emuToPx(blur).toFixed(1)}px)`);
    const alpha = isRecord(effects.alphaModulateFixed)
      ? num(effects.alphaModulateFixed.amount)
      : undefined;
    if (alpha != null && alpha < 100) opacity = Math.max(0, alpha / 100);
  }
  const shadow = outerShadowOf(pic.effects);
  if (!css.length && opacity == null && shadow == null) return undefined;
  return {
    ...(css.length ? { filter: css.join(" ") } : {}),
    ...(opacity != null ? { opacity } : {}),
    ...(shadow ? { shadow } : {}),
  };
}

/** Word's eight ST_RelativeHorizontalPosition values → the four semantic
 *  axes the painter resolves (margin/insideMargin/character → column,
 *  outsideMargin → rightMargin — the unmirrored reading). */
const H_RELATIVE: Record<string, LayoutDrawingAnchor["horizontal"]["relative"]> = {
  column: "column",
  margin: "column",
  insideMargin: "column",
  character: "column",
  leftMargin: "leftMargin",
  rightMargin: "rightMargin",
  outsideMargin: "rightMargin",
  page: "page",
};

/** ST_RelativeVerticalPosition → four axes (margin/insideMargin → topMargin,
 *  line → paragraph, outsideMargin → bottomMargin). */
const V_RELATIVE: Record<string, LayoutDrawingAnchor["vertical"]["relative"]> = {
  paragraph: "paragraph",
  line: "paragraph",
  margin: "topMargin",
  insideMargin: "topMargin",
  topMargin: "topMargin",
  bottomMargin: "bottomMargin",
  outsideMargin: "bottomMargin",
  page: "page",
};

/** One position axis (align > posOffset > percentOffset) → the anchor spec;
 *  an empty position collapses to offset 0 on the fallback axis. */
function anchorAxis<R extends string, A extends string>(
  pos: unknown,
  relativeTable: Record<string, R>,
  fallback: R,
  alignTable: Record<string, A>,
): { relative: R; offsetPx?: number; percent?: number; align?: A } {
  const axis = isRecord(pos) ? pos : {};
  const relative = relativeTable[str(axis.relative) ?? ""] ?? fallback;
  const align = alignTable[str(axis.align) ?? ""];
  if (align) return { relative, align };
  const offsetEmu = measureEmu(axis.offset);
  if (offsetEmu != null) return { relative, offsetPx: emuToPx(offsetEmu) };
  const pct = num(axis.percentOffset);
  if (pct != null) return { relative, percent: pct / 1000 };
  return { relative, offsetPx: 0 };
}

/** a:custGeom path coordinates arrive as strings (guide-resolved literals). */
function coord(v: unknown): number {
  const n = typeof v === "string" ? Number(v) : num(v);
  return n != null && Number.isFinite(n) ? n : 0;
}

/** a:custGeom pathLst → SVG path data scaled from the path's own space
 *  (path @w/@h) into the member box. moveTo/lineTo/quadBezTo/cubicBezTo/close
 *  convert directly; arcTo (elliptical-by-angle) emits cubic segments through
 *  the shared arc converter — the current point (tracked in box coordinates;
 *  close returns it to the subpath's start) is on the ellipse at the start
 *  angle, which is what the converter derives its center from. The command
 *  union is the parse contract, so a token mismatch fails at compile time. */
function customGeometryPath(
  cg: CustomGeometryOptions,
  width: number,
  height: number,
): string | undefined {
  const parts: string[] = [];
  const r2 = (v: number): number => Math.round(v * 100) / 100;
  for (const p of cg.pathList ?? []) {
    const sx = p.w ? width / p.w : 1;
    const sy = p.h ? height / p.h : 1;
    const x = (v: string): number => r2(coord(v) * sx);
    const y = (v: string): number => r2(coord(v) * sy);
    // Current point + subpath start in box coordinates (arcTo needs the
    // current point; close hands it back the subpath's start).
    let cx = 0;
    let cy = 0;
    let startX = 0;
    let startY = 0;
    for (const cmd of p.commands) {
      switch (cmd.command) {
        case "moveTo":
          cx = x(cmd.point.x);
          cy = y(cmd.point.y);
          startX = cx;
          startY = cy;
          parts.push(`M ${cx} ${cy}`);
          break;
        case "lineTo":
          cx = x(cmd.point.x);
          cy = y(cmd.point.y);
          parts.push(`L ${cx} ${cy}`);
          break;
        case "quadBezTo":
          parts.push(
            `Q ${x(cmd.points[0].x)} ${y(cmd.points[0].y)} ${x(cmd.points[1].x)} ${y(cmd.points[1].y)}`,
          );
          cx = x(cmd.points[1].x);
          cy = y(cmd.points[1].y);
          break;
        case "cubicBezTo":
          parts.push(
            `C ${x(cmd.points[0].x)} ${y(cmd.points[0].y)} ${x(cmd.points[1].x)} ${y(cmd.points[1].y)} ${x(cmd.points[2].x)} ${y(cmd.points[2].y)}`,
          );
          cx = x(cmd.points[2].x);
          cy = y(cmd.points[2].y);
          break;
        case "close":
          parts.push("Z");
          cx = startX;
          cy = startY;
          break;
        case "arcTo": {
          // Angles are 1/60000 degree (DrawingML's ST_Angle); radii scale
          // with their axis — an axis-aligned ellipse stays axis-aligned.
          const rx = coord(cmd.widthRadius) * sx;
          const ry = coord(cmd.heightRadius) * sy;
          const st = ((coord(cmd.startAngle) / 60000) * Math.PI) / 180;
          const sw = ((coord(cmd.sweepAngle) / 60000) * Math.PI) / 180;
          const segments = arcToSegments(cx, cy, rx, ry, st, sw);
          for (const seg of segments) {
            parts.push(
              `C ${r2(seg.c1x)} ${r2(seg.c1y)} ${r2(seg.c2x)} ${r2(seg.c2y)} ${r2(seg.x)} ${r2(seg.y)}`,
            );
          }
          const end = segments[segments.length - 1];
          if (end) {
            cx = end.x;
            cy = end.y;
          }
          break;
        }
      }
    }
  }
  return parts.length > 0 ? parts.join(" ") : undefined;
}

/** px-per-EMU for a group's child space: the box's px extent over chExt — or
 *  over the group's own EMU extent when chExt is absent (children share its
 *  units); no extent at all degrades to the plain EMU→px factor. */
function childScale(boxPx: number, chExt: number | undefined, extEmu: number | undefined): number {
  if (chExt) return boxPx / chExt;
  if (extEmu) return boxPx / extEmu;
  return emuToPx(1);
}

/** A flip on a group mirrors every descendant's box within that group's own
 *  box. Nested flips stack, so the recursion carries the list and applies
 *  each mirror outermost-first to the member's final box. */
interface GroupMirror {
  h: boolean;
  v: boolean;
  x: number;
  y: number;
  width: number;
  height: number;
}

/** One wps shape (group child data or a standalone WpsShapeOptions run) → the
 *  drawing members at `x,y` sized `width×height` (a multi-path preset yields
 *  one member per spec path — a shape's fill-only and stroke-only layers paint
 *  separately). A shape with txbx content is a text box: its paragraphs
 *  project as blocks (full style cascade) for the renderer to stack in the
 *  box. An empty children array is a shape without text, not an empty box. */
function wpsMemberOf(
  data: unknown,
  x: number,
  y: number,
  width: number,
  height: number,
  ctx: ProjectContext,
): LayoutDrawingMember[] {
  if (!isRecord(data)) return [];
  const fill = solidFillOf(data.fill);
  const line = outlineOf(data.outline);
  const shadow = outerShadowOf(data.effects);
  // 0.13.0 renamed ShapeCoreOptions.presetGeometry → geometry and widened it to
  // the ShapeType token shorthand | PresetGeometryOptions (whose
  // adjustmentValues override the preset's default guide values).
  const preset =
    typeof data.geometry === "string"
      ? data.geometry
      : isRecord(data.geometry)
        ? str(data.geometry.preset)
        : undefined;
  const adjustments =
    isRecord(data.geometry) && Array.isArray(data.geometry.adjustmentValues)
      ? (data.geometry.adjustmentValues as readonly GeometryGuide[])
      : undefined;
  const children = Array.isArray(data.children) ? data.children : [];
  // The shape's own a:xfrm @rot (degrees) — Word's diagonal watermark.
  // The xfrm flips stay unprojected here: the LayoutDrawing's flipH/flipV
  // (threaded below) hand them to the painter's single mirror group.
  const rotation = isRecord(data.transformation) ? num(data.transformation.rotation) : undefined;
  if (children.length > 0) {
    const bodyPr = isRecord(data.bodyProperties) ? data.bodyProperties : {};
    // Insets are EMU or universal measure; BODY_INSET_EMU is the Word
    // default applied whenever the side is absent.
    const ins = (v: unknown, fallback: number): number => {
      const emu = measureEmu(v);
      return emu != null ? emuToPx(emu) : emuToPx(fallback);
    };
    const blocks: LayoutParagraph[] = [];
    for (const p of children) {
      const block = projectParagraph(p as BodyParagraph, ctx);
      if (block) blocks.push(block);
    }
    // A non-box preset paints its evaluated silhouette instead of the plain
    // rectangle (Word's prstGeom under the body). The fill layers merge into
    // one d (a donut's subpaths share the fill); with none, the stroke-only
    // preset carries its outline path (line, arc). A straight line skips the
    // evaluator — its d is the box diagonal.
    const straight = preset === "line" || preset === "straightConnector1";
    const outlines = straight
      ? [{ d: linePathData(width, height), fill: false, stroke: true }]
      : preset && preset !== "rect" && preset !== "roundRect" && preset !== "ellipse"
        ? presetShapePaths(preset, width, height, adjustments)
        : undefined;
    const silhouette = outlines
      ? (outlines.some((o) => o.fill) ? outlines.filter((o) => o.fill) : outlines)
          .map((o) => o.d)
          .join(" ")
      : undefined;
    // A straight line's arrows ride beside the text member (they fill the
    // line color, so they cannot merge into the silhouette's paint) at the
    // unflipped diagonal's corners — the painter's mirror group flips them.
    const arrows = straight ? lineEndMembersOf(line, 0, 0, width, height) : [];
    return [
      {
        kind: "textBox",
        x,
        y,
        width,
        height,
        ...(rotation != null && rotation !== 0 ? { rotation } : {}),
        // The shape's own spPr paint — a txbx box draws under its text even
        // when the body is empty (Word's plain text box). The preset travels
        // with it: a text-carrying ellipse paints as an ellipse.
        ...(preset ? { preset } : {}),
        ...(silhouette ? { d: silhouette } : {}),
        ...(fill ? { fill } : {}),
        ...(line
          ? {
              line: {
                px: line.px,
                ...(line.color ? { color: line.color } : {}),
                ...(line.cap ? { cap: line.cap } : {}),
                ...(line.join ? { join: line.join } : {}),
                ...(line.dash ? { dash: line.dash } : {}),
              },
            }
          : {}),
        insets: (() => {
          // Word stacks the text inside the preset's text rectangle (the ECMA
          // "rect" guides — a circle keeps its words off the rim); the bodyPr
          // insets apply on top of that shrink. Presets without a rect
          // definition (the plain box) evaluate to undefined: zero extra.
          const tr = preset ? presetShapeTextRect(preset, width, height, adjustments) : undefined;
          return {
            left: ins(bodyPr.lIns, BODY_INSET_EMU.left) + (tr ? Math.max(0, tr.l) : 0),
            top: ins(bodyPr.tIns, BODY_INSET_EMU.top) + (tr ? Math.max(0, tr.t) : 0),
            right: ins(bodyPr.rIns, BODY_INSET_EMU.right) + (tr ? Math.max(0, width - tr.r) : 0),
            bottom: ins(bodyPr.bIns, BODY_INSET_EMU.bottom) + (tr ? Math.max(0, height - tr.b) : 0),
          };
        })(),
        // VerticalAnchor is already full-word ("top"/"center"/"bottom");
        // justify/distribute stretch to the box — treated as top until then.
        anchor: bodyPr.anchor === "center" || bodyPr.anchor === "bottom" ? bodyPr.anchor : "top",
        // bodyPr @vert — only the two rotated layouts project (the stacked
        // variants need per-glyph upright layout the renderer has no model
        // for); they render as horizontal, a registered gap.
        ...(bodyPr.vertical === "vertical" || bodyPr.vertical === "vertical270"
          ? { textVertical: bodyPr.vertical }
          : {}),
        ...(shadow ? { shadow } : {}),
        // a:spAutoFit: Word draws the box shrunk to its text — the declared
        // extent's height is stale and must not drive vertical centering.
        ...(bodyPr.spAutoFit === true ? { autoFit: true } : {}),
        // bodyPr @compatLnSpc is deliberately not threaded: Word's own layout
        // engine ignores it for wps text boxes (the txbxContent is laid out by
        // the standard paragraph rules — grid snap and half-leading included;
        // pixel-verified against the reference render), so the attribute only
        // matters to PowerPoint-native consumers.
        blocks,
      },
      ...arrows,
    ];
  }
  // Straight connector (a straight line across its box) and custom
  // geometry both project to path members; the box-like presets stay
  // shape members.
  if (preset === "line" || preset === "straightConnector1") {
    return [
      {
        kind: "path",
        x,
        y,
        width,
        height,
        d: linePathData(width, height),
        fill,
        line,
        ...(shadow ? { shadow } : {}),
      },
      // The unflipped diagonal's corners — the painter's mirror group flips
      // the whole member set (segment and arrows alike) in one group.
      ...lineEndMembersOf(line, 0, 0, width, height),
    ];
  }
  if (preset == null) {
    const d = data.customGeometry
      ? customGeometryPath(data.customGeometry as CustomGeometryOptions, width, height)
      : undefined;
    if (d)
      return [{ kind: "path", x, y, width, height, d, fill, line, ...(shadow ? { shadow } : {}) }];
    return [];
  }
  // Every non-box preset expands through the ECMA-376 evaluator into one
  // path member per spec path — the painter has no native geometry for
  // them, and a preset's fill-only/stroke-only layers must paint
  // separately (fill "none" paths carry no fill; stroke=false paths carry
  // no outline). An unknown token falls through to the shape member —
  // today's behavior for tokens the data module doesn't define.
  if (preset !== "rect" && preset !== "roundRect" && preset !== "ellipse") {
    const outlines = presetShapePaths(preset, width, height, adjustments);
    if (outlines) {
      return outlines.map((o) => ({
        kind: "path" as const,
        x,
        y,
        width,
        height,
        d: o.d,
        ...(o.fill && fill ? { fill } : {}),
        ...(o.stroke && line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
      }));
    }
  }
  const opacity = fillOpacityOf(data.fill);
  return [
    {
      kind: "shape",
      x,
      y,
      width,
      height,
      preset,
      fill,
      ...(opacity != null ? { opacity } : {}),
      line,
      ...(shadow ? { shadow } : {}),
    },
  ];
}

/** One group level's child-space → drawing-box-px mapping, threaded through
 *  the recursion: a member at child-space EMU `off` lands at
 *  `origin + (off - chOff) * scale` px; a nested group composes its own
 *  chOff/chExt on top (origin = its own box position, scale = its box extent
 *  over its chExt). Children are the office-open GroupChildMediaData union —
 *  the same contract stringify consumes, so field/token drift fails here at
 *  compile time. */
/** a:srcRect crops the source image per side, as signed fractions — negative
 *  insets (ST_Percentage < 0) pad the source outward. office-open's picture
 *  parse (readSourceRectangle) emits the RAW ST_Percentage int (100000 =
 *  100%), despite SourceRectangleOptions documenting integer percent — flip
 *  to /100 when that contract breach is fixed upstream. */
export function cropOf(
  pic: unknown,
): { left: number; top: number; right: number; bottom: number } | undefined {
  const sr = isRecord(pic) && isRecord(pic.sourceRectangle) ? pic.sourceRectangle : undefined;
  if (!sr) return undefined;
  const pct = (v: unknown): number | undefined =>
    typeof v === "number" && v !== 0 ? v / 100000 : undefined;
  const crop = {
    left: pct(sr.left) ?? 0,
    top: pct(sr.top) ?? 0,
    right: pct(sr.right) ?? 0,
    bottom: pct(sr.bottom) ?? 0,
  };
  return crop.left !== 0 || crop.top !== 0 || crop.right !== 0 || crop.bottom !== 0
    ? crop
    : undefined;
}

function walkGroup(
  group: { children: readonly GroupChildMediaData[] },
  originX: number,
  originY: number,
  scaleX: number,
  scaleY: number,
  chOffX: number,
  chOffY: number,
  out: LayoutDrawingMember[],
  ctx: ProjectContext,
  mirrors?: readonly GroupMirror[],
  path: readonly number[] = [],
): void {
  for (let i = 0; i < group.children.length; i++) {
    const child = group.children[i]!;
    const t: MediaDataTransformation = child.transformation;
    const off = t.offset?.emus;
    if (!off) continue;
    let x = originX + (off.x - chOffX) * scaleX;
    let y = originY + (off.y - chOffY) * scaleY;
    const width = t.emus.x * scaleX;
    const height = t.emus.y * scaleY;
    for (const m of mirrors ?? []) {
      if (m.h) x = 2 * m.x + m.width - x - width;
      if (m.v) y = 2 * m.y + m.height - y - height;
    }
    // The member's index path through the group children (nested groups
    // extend it) — how the PM side re-finds the child node from a member hit.
    const childPath = [...path, i];

    // Nested wpg group: flatten in place — its members land in this drawing's
    // box through the composed mapping (Word renders the group tree unrolled).
    if (child.type === "wpg") {
      const own: GroupMirror | undefined =
        t.flipHorizontal === true || t.flipVertical === true
          ? {
              h: t.flipHorizontal === true,
              v: t.flipVertical === true,
              x,
              y,
              width,
              height,
            }
          : undefined;
      walkGroup(
        child,
        x,
        y,
        childScale(width, child.childExtentWidth, t.emus.x),
        childScale(height, child.childExtentHeight, t.emus.y),
        child.childOffsetX ?? 0,
        child.childOffsetY ?? 0,
        out,
        ctx,
        own ? [...(mirrors ?? []), own] : mirrors,
        childPath,
      );
      continue;
    }

    if (child.type === "wps") {
      // Published 0.12.3 parse bug stringified nested shape data — a
      // non-object data skips the member (absence over corrupt geometry).
      if (child.data == null || typeof child.data !== "object") continue;
      for (const member of wpsMemberOf(child.data, x, y, width, height, ctx)) {
        out.push({ ...member, childPath });
      }
    } else if (child.type === "chart") {
      // The group child carries the bare ChartSpaceOptions (chartOptions) —
      // the transformation lives on the member itself.
      const data = (child as { chartOptions?: unknown }).chartOptions;
      if (data != null && typeof data === "object") {
        out.push({ kind: "chart", x, y, width, height, childPath, chart: data });
      }
    } else {
      // Everything else is treated as a picture member: real media children
      // carry bytes; chart/contentPart children have none and pictureSrc
      // yields undefined — the painter's empty-frame placeholder. A metafile
      // child expands into its vector replay instead, offset by the child
      // box (replay members are box-relative).
      const replay = metafileMembers(child, width, height, cropOf(child));
      if (replay) {
        out.push(...replay.map((m) => ({ ...m, x: m.x + x, y: m.y + y, childPath })));
      } else {
        out.push({
          kind: "picture",
          x,
          y,
          width,
          height,
          childPath,
          src: pictureSrc(child),
          flipH: t.flipHorizontal === true || undefined,
          flipV: t.flipVertical === true || undefined,
          crop: cropOf(child),
        });
      }
    }
  }
}

/** One wpg group run (GroupOptions) → a LayoutDrawing anchored to its
 *  paragraph. Members carry the group's child coordinate space (chOff/chExt)
 *  already resolved into px-in-box, nested groups flattened. A wps child
 *  whose `data` is not a record is skipped — the published 0.12.3 parse
 *  stringified nested shape data, so those members render as absence rather
 *  than corrupt geometry. */
function projectDrawing(group: GroupOptions, ctx: ProjectContext): LayoutDrawing | undefined {
  const extW = measureEmu(group.transformation.width);
  const extH = measureEmu(group.transformation.height);
  if (extW == null || extH == null || extW <= 0 || extH <= 0) return undefined;
  const rotation = group.transformation.rotation;
  const { anchor, wrap, wrapSide, contour, behind, zIndex, distances } = drawingAnchorOf(
    group.floating,
    emuToPx(extW),
    emuToPx(extH),
  );

  // Child coordinate space: chOff/chExt → the group's EMU box. A missing
  // chExt means the children already live in the group's own units (1:1).
  // A flip on the top-level group mirrors that whole child space within the
  // drawing's own box — the same mirror stack nested groups extend.
  const topMirror: GroupMirror | undefined =
    group.transformation.flipHorizontal === true || group.transformation.flipVertical === true
      ? {
          h: group.transformation.flipHorizontal === true,
          v: group.transformation.flipVertical === true,
          x: 0,
          y: 0,
          width: emuToPx(extW),
          height: emuToPx(extH),
        }
      : undefined;
  const members: LayoutDrawingMember[] = [];
  walkGroup(
    group,
    0,
    0,
    childScale(emuToPx(extW), group.childExtentWidth, extW),
    childScale(emuToPx(extH), group.childExtentHeight, extH),
    group.childOffsetX ?? 0,
    group.childOffsetY ?? 0,
    members,
    ctx,
    topMirror ? [topMirror] : undefined,
  );
  return {
    anchor,
    width: emuToPx(extW),
    height: emuToPx(extH),
    members,
    wrap,
    wrapSide,
    ...(contour ? { contour } : {}),
    behind,
    ...(zIndex != null ? { zIndex } : {}),
    distances,
    ...(rotation ? { rotation } : {}),
  };
}

/** wp:anchor positioning shared by every floating drawing kind (group, wps
 *  shape, picture) — every relativeFrom axis plus the offset/align choice;
 *  the painter owns the page geometry each axis resolves against. Wrap modes
 *  that keep the box out of the text flow (none, through's transparent
 *  interior) map to undefined. The wrap distances (w:anchor distL/T/R/B,
 *  floating.margins) thread through: zones and bands pad by them. The tight/
 *  through contour polygon scales out of Word's 21600×21600 wrap space onto
 *  the px extent (`widthPx`/`heightPx`, box-relative). */
function drawingAnchorOf(
  floating: unknown,
  widthPx = 0,
  heightPx = 0,
): {
  anchor: LayoutDrawingAnchor;
  wrap: "square" | "tight" | "topAndBottom" | undefined;
  wrapSide: LayoutDrawing["wrapSide"];
  contour: LayoutDrawing["contour"];
  behind: boolean | undefined;
  zIndex: number | undefined;
  distances: LayoutDrawing["distances"];
} {
  const f = isRecord(floating) ? floating : {};
  const anchor: LayoutDrawingAnchor = {
    horizontal: anchorAxis(f.horizontalPosition, H_RELATIVE, "column" as const, {
      left: "left",
      inside: "left",
      center: "center",
      right: "right",
      outside: "right",
    }),
    vertical: anchorAxis(f.verticalPosition, V_RELATIVE, "paragraph" as const, {
      top: "top",
      inside: "top",
      center: "center",
      bottom: "bottom",
      outside: "bottom",
    }),
  };
  const wrapType = isRecord(f.wrap) ? f.wrap.type : undefined;
  const wrap =
    wrapType === "square" || wrapType === "through"
      ? ("square" as const)
      : wrapType === "tight"
        ? ("tight" as const)
        : wrapType === "topAndBottom"
          ? ("topAndBottom" as const)
          : undefined;
  // ST_WrapSide: which side of the box takes text (square/tight only).
  const rawSide = isRecord(f.wrap) ? str(f.wrap.side) : undefined;
  const wrapSide =
    rawSide === "left" || rawSide === "right" || rawSide === "largest"
      ? rawSide
      : rawSide === "bothSides"
        ? ("both" as const)
        : undefined;
  // The wrapPolygon's points live in Word's 21600×21600 space, stretched
  // onto the extent box per axis (LibreOffice's GraphicImport does the same).
  const polygon =
    isRecord(f.wrap) && isRecord(f.wrap.polygon) && Array.isArray(f.wrap.polygon.points)
      ? f.wrap.polygon.points
      : undefined;
  const contour =
    polygon && polygon.length >= 3 && widthPx > 0 && heightPx > 0
      ? polygon
          .filter((p: unknown) => isRecord(p))
          .map((p: Rec) => ({
            x: ((num(p.x) ?? 0) / 21600) * widthPx,
            y: ((num(p.y) ?? 0) / 21600) * heightPx,
          }))
      : undefined;
  // Wrap distances: EMU (or a UniversalMeasure) per side → px. wrapNone never
  // reads them, but carrying them costs nothing and keeps round-trips honest.
  const margins = isRecord(f.margins) ? f.margins : undefined;
  const distPx = (v: unknown): number | undefined => {
    const emu = measureEmu(v);
    return emu != null ? emuToPx(emu) : undefined;
  };
  const distances =
    margins &&
    (margins.left != null || margins.top != null || margins.right != null || margins.bottom != null)
      ? {
          left: distPx(margins.left),
          top: distPx(margins.top),
          right: distPx(margins.right),
          bottom: distPx(margins.bottom),
        }
      : undefined;
  return {
    anchor,
    wrap,
    wrapSide,
    contour,
    // Word 2013+ honors behindDoc for wrapNone anchors only: a wrapped box
    // (square/tight/through/topAndBottom) always paints opaque in front of
    // the text, regardless of the attribute.
    behind: wrap == null ? f.behindDocument === true || undefined : undefined,
    zIndex: typeof f.zIndex === "number" ? f.zIndex : undefined,
    distances,
  };
}

/** A standalone floating picture run (wp:anchor pic:pic, PictureOptions):
 *  one drawing whose single member is the image filling its own box. */
function projectFloatingPicture(pic: Rec): LayoutDrawing | undefined {
  const tr = isRecord(pic.transformation) ? pic.transformation : {};
  const w = measureEmu(tr.width);
  const h = measureEmu(tr.height);
  if (w == null || h == null || w <= 0 || h <= 0) return undefined;
  const width = emuToPx(w);
  const height = emuToPx(h);
  const { anchor, wrap, wrapSide, contour, behind, zIndex, distances } = drawingAnchorOf(
    pic.floating,
    width,
    height,
  );
  const crop = cropOf(pic);
  const adjust = pictureAdjustOf(pic);
  const line = outlineOf(pic.outline);
  return {
    anchor,
    width,
    height,
    wrap,
    wrapSide,
    ...(contour ? { contour } : {}),
    behind,
    ...(zIndex != null ? { zIndex } : {}),
    distances,
    ...(typeof tr.rotation === "number" && tr.rotation ? { rotation: tr.rotation } : {}),
    ...(tr.flipHorizontal === true ? { flipH: true } : {}),
    ...(tr.flipVertical === true ? { flipV: true } : {}),
    // A srcRect-cropped metafile replay reaches past the extent — flag it so
    // the painter clips (GDI playback semantics); the flat member never does.
    ...(crop ? { clipMembers: true } : {}),
    // A metafile picture expands into its vector replay (the srcRect crop
    // folds into the replay's frame mapping); anything else stays one flat
    // member with the crop on the raster source.
    members: metafileMembers(pic, width, height, crop) ?? [
      {
        kind: "picture",
        x: 0,
        y: 0,
        width,
        height,
        src: pictureSrc(pic as { type?: unknown; data?: unknown }),
        crop,
        ...(adjust?.filter ? { filter: adjust.filter } : {}),
        ...(adjust?.opacity != null ? { opacity: adjust.opacity } : {}),
        ...(adjust?.shadow ? { shadow: adjust.shadow } : {}),
        ...(line ? { line } : {}),
      },
    ],
  };
}

/** A standalone floating wps shape run (WpsShapeOptions): the same member
 *  projection a wps child inside a wpg group gets, anchored to the
 *  paragraph in its own one-member drawing. */
function projectWpsShapeRun(wps: Rec, ctx: ProjectContext): LayoutDrawing | undefined {
  const tr = isRecord(wps.transformation) ? wps.transformation : {};
  const w = measureEmu(tr.width);
  const h = measureEmu(tr.height);
  if (w == null || h == null || w <= 0 || h <= 0) return undefined;
  const members = wpsMemberOf(wps, 0, 0, emuToPx(w), emuToPx(h), ctx);
  if (members.length === 0) return undefined;
  const { anchor, wrap, wrapSide, contour, behind, zIndex, distances } = drawingAnchorOf(
    wps.floating,
    emuToPx(w),
    emuToPx(h),
  );
  return {
    anchor,
    width: emuToPx(w),
    height: emuToPx(h),
    members,
    wrap,
    wrapSide,
    ...(contour ? { contour } : {}),
    behind,
    ...(zIndex != null ? { zIndex } : {}),
    distances,
    ...(typeof tr.rotation === "number" && tr.rotation ? { rotation: tr.rotation } : {}),
    ...(tr.flipHorizontal === true ? { flipH: true } : {}),
    ...(tr.flipVertical === true ? { flipV: true } : {}),
  };
}

/** One chart run (ChartOptions flattened on attrs.chart) → a single-member
 *  drawing. The chart paints inside its extent box; the anchor spec resolves
 *  exactly like a shape's (floating) or stays inline (the inline-picture
 *  path handles that — here only the floating arm projects). */
function projectChartRun(chart: Rec): LayoutDrawing | undefined {
  const tr = isRecord(chart.transformation) ? chart.transformation : {};
  const w = measureEmu(tr.width);
  const h = measureEmu(tr.height);
  if (w == null || h == null || w <= 0 || h <= 0) return undefined;
  if (!isRecord(chart.floating)) return undefined;
  const { anchor, wrap, wrapSide, contour, behind, zIndex, distances } = drawingAnchorOf(
    chart.floating,
    emuToPx(w),
    emuToPx(h),
  );
  return {
    anchor,
    width: emuToPx(w),
    height: emuToPx(h),
    members: [{ kind: "chart", x: 0, y: 0, width: emuToPx(w), height: emuToPx(h), chart }],
    wrap,
    wrapSide,
    ...(contour ? { contour } : {}),
    behind,
    ...(zIndex != null ? { zIndex } : {}),
    distances,
    ...(typeof tr.rotation === "number" && tr.rotation ? { rotation: tr.rotation } : {}),
    ...(tr.flipHorizontal === true ? { flipH: true } : {}),
    ...(tr.flipVertical === true ? { flipV: true } : {}),
  };
}

/** Collect the anchored drawing runs of one paragraph (top level and one
 *  nested run level — a drawing rides its own w:r): wpg groups, wps shapes,
 *  charts, and floating pictures. Non-floating pictures stay inline atoms. */
export function projectDrawings(runs: readonly unknown[], ctx: ProjectContext): LayoutDrawing[] {
  const out: LayoutDrawing[] = [];
  const each = (run: Rec): void => {
    if (isRecord(run.wpgGroup)) {
      const d = projectDrawing(run.wpgGroup as unknown as GroupOptions, ctx);
      if (d) out.push(d);
    }
    if (isRecord(run.wpsShape)) {
      const d = projectWpsShapeRun(run.wpsShape, ctx);
      if (d) out.push(d);
    }
    if (isRecord(run.chart)) {
      const d = projectChartRun(run.chart);
      if (d) out.push(d);
    }
    if (isRecord(run.picture) && isRecord(run.picture.floating)) {
      const d = projectFloatingPicture(run.picture);
      if (d) out.push(d);
    }
  };
  for (const run of runs) {
    if (!isRecord(run)) continue;
    each(run);
    if (Array.isArray(run.children)) {
      for (const inner of run.children) if (isRecord(inner)) each(inner);
    }
  }
  return out;
}
