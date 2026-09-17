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
  const x = t.sx * emuOf(group.x) + t.dx;
  const y = t.sy * emuOf(group.y) + t.dy;
  const w = t.sx * emuOf(group.width);
  const h = t.sy * emuOf(group.height);
  // An unset chOff/chExt stringifies to the group box itself (identity).
  const chW = emuOf(group.childExtentWidth);
  const chH = emuOf(group.childExtentHeight);
  const sx = chW > 0 ? w / chW : 1;
  const sy = chH > 0 ? h / chH : 1;
  const inner: Xform = {
    sx,
    sy,
    dx: x - emuOf(group.childOffsetX) * sx,
    dy: y - emuOf(group.childOffsetY) * sy,
  };
  return childMembers(group.children ?? [], inner, path);
}
