// The children walk: a slide's (or group's) children in document order —
// group chains recurse through their chOff/chExt affine map, every other
// child dispatches to its member domain.

import type { LayoutDrawingMember } from "@docen/layout";
import type { GroupOptions, SlideChild } from "@office-open/pptx";

import { IDENTITY, emuOf, geometryOf, type Xform } from "./geometry";
import { endpointMembers } from "./lines";
import { pictureMember } from "./pictures";
import { shapeMembers } from "./shapes";
import { tableMember } from "./tables";

export function childMembers(
  children: readonly SlideChild[],
  t: Xform,
  prefix: readonly number[],
): LayoutDrawingMember[] {
  const out: LayoutDrawingMember[] = [];
  children.forEach((child, i) => {
    // The group-address path only exists below a group — a slide's own
    // members have nothing above them to address into.
    const cp = prefix.length > 0 ? [...prefix, i] : undefined;
    if ("shape" in child) {
      out.push(...shapeMembers(child.shape, t, cp));
    } else if ("picture" in child) {
      const m = pictureMember(child.picture, t, cp);
      if (m) out.push(m);
    } else if ("line" in child) {
      out.push(...endpointMembers(child.line, t, cp, {}));
    } else if ("connector" in child) {
      out.push(
        ...endpointMembers(
          child.connector,
          t,
          cp,
          geometryOf(child.connector.properties?.geometry),
        ),
      );
    } else if ("group" in child) {
      out.push(...groupMembers(child.group, t, [...prefix, i]));
    } else if ("table" in child) {
      out.push(tableMember(child.table, t, cp));
    } else if ("chart" in child) {
      const c = child.chart;
      out.push({
        kind: "chart",
        x: t.sx * emuOf(c.x) + t.dx,
        y: t.sy * emuOf(c.y) + t.dy,
        width: t.sx * emuOf(c.width),
        height: t.sy * emuOf(c.height),
        ...(cp ? { childPath: cp } : {}),
        chart: c,
      });
    }
  });
  return out;
}

function groupMembers(
  group: GroupOptions,
  t: Xform,
  path: readonly number[],
): LayoutDrawingMember[] {
  return childMembers(group.children ?? [], groupXformOf(group, t), path);
}

function groupXformOf(group: GroupOptions, t: Xform): Xform {
  const x = t.sx * emuOf(group.x) + t.dx;
  const y = t.sy * emuOf(group.y) + t.dy;
  const w = t.sx * emuOf(group.width);
  const h = t.sy * emuOf(group.height);
  // An unset chOff/chExt stringifies to the group box itself (identity).
  const chW = emuOf(group.childExtentWidth);
  const chH = emuOf(group.childExtentHeight);
  const sx = chW > 0 ? w / chW : 1;
  const sy = chH > 0 ? h / chH : 1;
  return {
    sx,
    sy,
    dx: x - emuOf(group.childOffsetX) * sx,
    dy: y - emuOf(group.childOffsetY) * sy,
  };
}

// ── group member hits ──

/** A group-nested member: the child-index chain under a group child, with
 *  its slide-absolute box and the child-space → px scale its parent maps it
 *  through (drag deltas divide by it). */
export interface MemberHit {
  path: number[];
  x: number;
  y: number;
  width: number;
  height: number;
  scale: { sx: number; sy: number };
}

const contains = (
  b: { x: number; y: number; width: number; height: number },
  x: number,
  y: number,
): boolean => x >= b.x && x <= b.x + b.width && y >= b.y && y <= b.y + b.height;

/** The child's outer box under the affine map (slide-absolute px) — the
 *  transform variants from their xfrm, lines and connectors from their
 *  endpoints. Unplaced children hit nothing. */
function childBoxUnder(
  child: SlideChild,
  t: Xform,
): { x: number; y: number; width: number; height: number } | null {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    const x1 = t.sx * emuOf(o.x1) + t.dx;
    const y1 = t.sy * emuOf(o.y1) + t.dy;
    const x2 = t.sx * emuOf(o.x2) + t.dx;
    const y2 = t.sy * emuOf(o.y2) + t.dy;
    return {
      x: Math.min(x1, x2),
      y: Math.min(y1, y2),
      width: Math.abs(x2 - x1),
      height: Math.abs(y2 - y1),
    };
  }
  const g =
    "shape" in child
      ? child.shape
      : "picture" in child
        ? child.picture
        : "group" in child
          ? child.group
          : "table" in child
            ? child.table
            : "chart" in child
              ? child.chart
              : undefined;
  if (!g) return null;
  return {
    x: t.sx * emuOf(g.x) + t.dx,
    y: t.sy * emuOf(g.y) + t.dy,
    width: t.sx * emuOf(g.width),
    height: t.sy * emuOf(g.height),
  };
}

const hitOf = (
  path: number[],
  box: { x: number; y: number; width: number; height: number },
  t: Xform,
): MemberHit => ({ path, ...box, scale: { sx: t.sx, sy: t.sy } });

/** The deepest group-nested member whose box contains the slide-local point
 *  — a child-index chain under the group child — or null: the point misses
 *  every member and the group itself is the hit. A nested group whose box
 *  contains the point but no member of its own does reads as the group. */
export function memberAt(child: SlideChild, x: number, y: number): MemberHit | null {
  if (!("group" in child)) return null;
  return hitMemberIn(child.group.children ?? [], groupXformOf(child.group, IDENTITY), x, y, []);
}

function hitMemberIn(
  children: readonly SlideChild[],
  t: Xform,
  x: number,
  y: number,
  prefix: number[],
): MemberHit | null {
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i]!;
    const box = childBoxUnder(c, t);
    if (!box || !contains(box, x, y)) continue;
    if ("group" in c) {
      const deeper = hitMemberIn(c.group.children ?? [], groupXformOf(c.group, t), x, y, [
        ...prefix,
        i,
      ]);
      if (deeper) return deeper;
    }
    return hitOf([...prefix, i], box, t);
  }
  return null;
}

/** The member at `path` (child indexes under the group child) — its box and
 *  scale, or null when the path dangles (the deck changed under the
 *  selection). */
export function memberByPath(child: SlideChild, path: readonly number[]): MemberHit | null {
  if (!("group" in child) || path.length === 0) return null;
  return memberInPath(child.group.children ?? [], groupXformOf(child.group, IDENTITY), path, 0);
}

function memberInPath(
  children: readonly SlideChild[],
  t: Xform,
  path: readonly number[],
  depth: number,
): MemberHit | null {
  const c = children[path[depth]!];
  if (!c) return null;
  const box = childBoxUnder(c, t);
  if (depth === path.length - 1) return box ? hitOf([...path], box, t) : null;
  if (!("group" in c)) return null;
  return memberInPath(c.group.children ?? [], groupXformOf(c.group, t), path, depth + 1);
}
