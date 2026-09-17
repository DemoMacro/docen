// Shape (p:sp) projection: box presets paint the plain shape member, straight
// presets the box diagonal, every other preset expands through the ECMA-376
// evaluator — and a text-carrying shape projects its body instead.

import {
  fillOpacityOf,
  lineEndMembersOf,
  linePathData,
  measureEmu,
  outlineOf,
  outerShadowOf,
  presetShapePaths,
  presetShapeTextRect,
  solidFillOf,
} from "@docen/core/geometry";
import {
  emuToPx,
  type LayoutDrawingLine,
  type LayoutDrawingMember,
  type LayoutDrawingShadow,
} from "@docen/layout";
import type { GeometryGuide } from "@office-open/core";
import type { TextBodyOptions } from "@office-open/core/drawing";
import type { ShapeOptions } from "@office-open/pptx";

import { BOX_PRESETS, STRAIGHT_PRESETS, emuOf, geometryOf, type Xform } from "./geometry";
import { textBlocks } from "./text";

// DrawingML bodyPr default insets (lIns/rIns 0.1", tIns/bIns 0.05").
const BODY_INSET_EMU = { left: 91440, top: 45720, right: 91440, bottom: 45720 };

interface ShapePaint {
  x: number;
  y: number;
  w: number;
  h: number;
  preset?: string;
  adjustments?: readonly GeometryGuide[];
  fill?: string;
  opacity?: number;
  line?: LayoutDrawingLine;
  shadow?: LayoutDrawingShadow;
  rotation?: number;
  childPath?: readonly number[];
}

export function shapeMembers(
  shape: ShapeOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
): LayoutDrawingMember[] {
  const paint: ShapePaint = {
    x: t.sx * emuOf(shape.x) + t.dx,
    y: t.sy * emuOf(shape.y) + t.dy,
    w: t.sx * emuOf(shape.width),
    h: t.sy * emuOf(shape.height),
    ...geometryOf(shape.properties?.geometry),
    fill: solidFillOf(shape.properties?.fill),
    opacity: fillOpacityOf(shape.properties?.fill),
    line: outlineOf(shape.properties?.outline),
    shadow: outerShadowOf(shape.properties?.effects),
    ...(shape.rotation ? { rotation: shape.rotation } : {}),
    ...(childPath ? { childPath } : {}),
  };
  // A text-carrying shape projects its body (PowerPoint writes an empty a:p
  // into every bare shape — only non-empty content makes a text box).
  if (hasText(shape.textBody)) {
    return [textBoxMember(shape.textBody, paint)];
  }
  const {
    x,
    y,
    w,
    h,
    preset,
    adjustments,
    fill,
    opacity,
    line,
    shadow,
    rotation,
    childPath: cp,
  } = paint;

  // Box presets (or none) paint as the plain shape member the renderer knows.
  if (!preset || BOX_PRESETS.has(preset)) {
    return [
      {
        kind: "shape",
        x,
        y,
        width: w,
        height: h,
        ...(preset && preset !== "rect" ? { preset } : {}),
        ...(fill ? { fill } : {}),
        ...(opacity != null ? { opacity } : {}),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
    ];
  }
  // A straight line paints the box diagonal; the stroke's end arrows expand
  // into their own fill members beside it.
  if (STRAIGHT_PRESETS.has(preset)) {
    return [
      {
        kind: "path",
        x,
        y,
        width: w,
        height: h,
        d: linePathData(w, h),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
      ...lineEndMembersOf(line, 0, 0, w, h).map((m) => ({
        ...m,
        x: m.x + x,
        y: m.y + y,
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      })),
    ];
  }
  // Every other preset expands through the ECMA-376 evaluator — one path
  // member per spec path (fill-only/stroke-only layers paint separately).
  // An unknown token falls back to the plain rectangle.
  const outlines = presetShapePaths(preset, w, h, adjustments);
  if (!outlines) {
    return [
      {
        kind: "shape",
        x,
        y,
        width: w,
        height: h,
        ...(fill ? { fill } : {}),
        ...(opacity != null ? { opacity } : {}),
        ...(line ? { line } : {}),
        ...(shadow ? { shadow } : {}),
        ...(rotation ? { rotation } : {}),
        ...(cp ? { childPath: cp } : {}),
      },
    ];
  }
  return outlines.map((o) => ({
    kind: "path",
    x,
    y,
    width: w,
    height: h,
    d: o.d,
    ...(o.fill && fill ? { fill } : {}),
    ...(o.fill && opacity != null ? { opacity } : {}),
    ...(o.stroke && line ? { line } : {}),
    ...(shadow ? { shadow } : {}),
    ...(rotation ? { rotation } : {}),
    ...(cp ? { childPath: cp } : {}),
  }));
}

/** A shape's body counts as text only with non-empty paragraph content —
 *  PowerPoint writes an empty a:p into every bare shape. */
function hasText(body: TextBodyOptions | undefined): body is TextBodyOptions {
  if (!body) return false;
  if (body.text) return true;
  return (body.paragraphs ?? []).some((p) =>
    typeof p === "string" ? p.length > 0 : p.text != null || (p.children?.length ?? 0) > 0,
  );
}

function textBoxMember(body: TextBodyOptions, paint: ShapePaint): LayoutDrawingMember {
  const { w, h, preset, adjustments } = paint;
  const straight = preset != null && STRAIGHT_PRESETS.has(preset);
  // A non-box preset paints its evaluated silhouette under the text (a text
  // ellipse stays an ellipse); the fill layers merge into one d. A straight
  // line skips the evaluator — its d is the box diagonal.
  const outlines = straight
    ? [{ d: linePathData(w, h), fill: false, stroke: true }]
    : preset && !BOX_PRESETS.has(preset)
      ? presetShapePaths(preset, w, h, adjustments)
      : undefined;
  const silhouette =
    outlines
      ?.filter((o) => o.fill)
      .map((o) => o.d)
      .join(" ") || undefined;
  const bp = body.bodyProperties;
  // The top-level sugar fields merge into bodyProperties on stringify (an
  // explicit bodyProperties field wins) — the projection reads the same
  // merge so authoring-shaped bodies anchor like their parsed form.
  const anchor = bp?.anchor ?? body.anchor;
  const ins = (v: unknown, def: number) => emuToPx(measureEmu(v) ?? def);
  // Text stacks inside the preset's text rectangle (a circle keeps its words
  // off the rim); the bodyPr insets apply on top of that shrink.
  const tr = preset && !straight ? presetShapeTextRect(preset, w, h, adjustments) : undefined;
  return {
    kind: "textBox",
    x: paint.x,
    y: paint.y,
    width: w,
    height: h,
    ...(preset ? { preset } : {}),
    ...(silhouette ? { d: silhouette } : {}),
    ...(paint.fill ? { fill: paint.fill } : {}),
    ...(paint.opacity != null ? { opacity: paint.opacity } : {}),
    ...(paint.line ? { line: paint.line } : {}),
    ...(paint.shadow ? { shadow: paint.shadow } : {}),
    ...(paint.childPath ? { childPath: paint.childPath } : {}),
    insets: {
      left: ins(bp?.lIns, BODY_INSET_EMU.left) + (tr ? Math.max(0, tr.l) : 0),
      top: ins(bp?.tIns, BODY_INSET_EMU.top) + (tr ? Math.max(0, tr.t) : 0),
      right: ins(bp?.rIns, BODY_INSET_EMU.right) + (tr ? Math.max(0, w - tr.r) : 0),
      bottom: ins(bp?.bIns, BODY_INSET_EMU.bottom) + (tr ? Math.max(0, h - tr.b) : 0),
    },
    anchor: anchor === "center" || anchor === "bottom" ? anchor : "top",
    ...(bp?.vertical === "vertical" || bp?.vertical === "vertical270"
      ? { textVertical: bp.vertical }
      : {}),
    ...(bp?.spAutoFit === true || body.autoFit === "shape" ? { autoFit: true } : {}),
    blocks: textBlocks(body),
  };
}
