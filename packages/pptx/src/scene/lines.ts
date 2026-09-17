// Line (p:sp) and connector (p:cxnSp) projection: endpoints to a box-diagonal
// path (direction carried), bent presets expanded through the evaluator.

import {
  lineEndMembersOf,
  outlineOf,
  outerShadowOf,
  presetShapePaths,
  solidFillOf,
} from "@docen/core/geometry";
import type { LayoutDrawingMember } from "@docen/layout";
import type { GeometryGuide } from "@office-open/core";
import type { ConnectorOptions, LineShapeOptions } from "@office-open/pptx";

import { STRAIGHT_PRESETS, emuOf, type Xform } from "./geometry";

export function endpointMembers(
  o: LineShapeOptions | ConnectorOptions,
  t: Xform,
  childPath: readonly number[] | undefined,
  g: { preset?: string; adjustments?: readonly GeometryGuide[] },
): LayoutDrawingMember[] {
  const x1 = emuOf(o.x1);
  const y1 = emuOf(o.y1);
  const x2 = emuOf(o.x2);
  const y2 = emuOf(o.y2);
  const x = t.sx * Math.min(x1, x2) + t.dx;
  const y = t.sy * Math.min(y1, y2) + t.dy;
  const w = t.sx * Math.abs(x2 - x1);
  const h = t.sy * Math.abs(y2 - y1);
  const line = outlineOf(o.properties?.outline);
  const fill = solidFillOf(o.properties?.fill);
  const shadow = outerShadowOf(o.properties?.effects);
  const p = (v: number): string => String(Math.round(v * 100) / 100);
  // Endpoints carry the direction (parse resolved the xfrm flips into them):
  // the segment runs between the corners the endpoints sit on.
  const lx1 = x1 <= x2 ? 0 : w;
  const ly1 = y1 <= y2 ? 0 : h;
  const d = `M ${p(lx1)} ${p(ly1)} L ${p(w - lx1)} ${p(h - ly1)}`;
  const cp = () => (childPath ? { childPath } : {});

  // A bent/elbow connector expands through the evaluator — but an evaluated
  // path always runs (0,0)→(w,h), and a reversed endpoint pair would need a
  // mirror the member model has no slot for, so those stay the diagonal
  // (registered gap).
  const outlines =
    g.preset && !STRAIGHT_PRESETS.has(g.preset) && x1 <= x2 && y1 <= y2
      ? presetShapePaths(g.preset, w, h, g.adjustments)
      : undefined;
  if (outlines) {
    return outlines.map((part) => ({
      kind: "path",
      x,
      y,
      width: w,
      height: h,
      d: part.d,
      ...(part.fill && fill ? { fill } : {}),
      ...(part.stroke && line ? { line } : {}),
      ...(shadow ? { shadow } : {}),
      ...cp(),
    }));
  }
  return [
    {
      kind: "path",
      x,
      y,
      width: w,
      height: h,
      d,
      ...(line ? { line } : {}),
      ...(shadow ? { shadow } : {}),
      ...cp(),
    },
    ...lineEndMembersOf(line, lx1, ly1, w - lx1, h - ly1).map((m) => ({
      ...m,
      x: m.x + x,
      y: m.y + y,
      ...cp(),
    })),
  ];
}
