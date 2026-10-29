// The children walk: a slide's (or group's) children in document order —
// group chains recurse through their chOff/chExt affine map, every other
// child dispatches to its member domain.

import { resolveChartColors } from "@docen/core/paint";
import type { LayoutDrawingMember } from "@docen/layout";
import { EMU_PER_PX } from "@docen/layout";
import type { ColorSchemeOptions } from "@office-open/core/theme";
import type { GroupOptions, PictureOptions, SlideChild } from "@office-open/pptx";

import { IDENTITY, emuOf, geometryOf, type Xform } from "./geometry";
import { endpointMembers } from "./lines";
import { mediaMember } from "./media";
import { pictureMember } from "./pictures";
import { shapeMembers } from "./shapes";
import { smartArtMember } from "./smartart";
import { SCHEME_SLOTS, type ThemeColors } from "./table-style";
import { tableMember } from "./tables";
import type { TextFieldContext } from "./text";

/** A color-scheme slot's hex (a string value or a sysClr's lastClr). */
function schemeColorOf(value: string | { lastClr?: string } | undefined): string | undefined {
  if (value === undefined) return undefined;
  const hex = typeof value === "string" ? value : value.lastClr;
  return hex?.replace("#", "").toUpperCase();
}

/** Merge a chart-owned theme override over its owning master's scheme. */
function chartThemeColorsOf(
  base: ThemeColors | undefined,
  scheme: ColorSchemeOptions | undefined,
): ThemeColors | undefined {
  if (!scheme) return base;
  const overrides = SCHEME_SLOTS.flatMap((key) => {
    const color = schemeColorOf(scheme[key]);
    return color ? [[key, color] as const] : [];
  });
  return overrides.length > 0 ? { ...base, ...Object.fromEntries(overrides) } : base;
}

/** The child's cNvPr `@hidden`, including graphic-frame chart/table surfaces.
 *  Hidden objects drop out of every projection. */
function hiddenOf(child: SlideChild): boolean {
  const nv =
    "shape" in child
      ? child.shape
      : "picture" in child
        ? child.picture
        : "line" in child
          ? child.line
          : "connector" in child
            ? child.connector
            : "group" in child
              ? child.group
              : "smartart" in child
                ? child.smartart
                : "video" in child
                  ? child.video
                  : "audio" in child
                    ? child.audio
                    : "ole" in child
                      ? child.ole
                      : "chart" in child
                        ? child.chart
                        : "table" in child
                          ? child.table
                          : undefined;
  return nv?.hidden === true;
}

/** A branch the projection can actually paint. PowerPoint's mc:Choice may
 *  lead with an opaque extension; a later Choice or the Fallback is often the
 *  standard shape/table/chart the editor can render. */
function projectableChildOf(child: SlideChild): boolean {
  if ("contentPart" in child || "rawXml" in child) return false;
  if ("alternateContent" in child)
    return (
      child.alternateContent.choices.some(({ child: branch }) => projectableChildOf(branch)) ||
      child.alternateContent.fallback !== undefined
    );
  return true;
}

function alternateChildOf(
  alternate: Extract<SlideChild, { alternateContent: unknown }>,
): SlideChild | undefined {
  return (
    alternate.alternateContent.choices.map(({ child }) => child).find(projectableChildOf) ??
    alternate.alternateContent.fallback
  );
}

export function childMembers(
  children: readonly SlideChild[],
  t: Xform,
  prefix: readonly number[],
  context: TextFieldContext = {},
): ProjectedSlideMember[] {
  const out: ProjectedSlideMember[] = [];
  children.forEach((child, i) => {
    if (hiddenOf(child)) return;
    const start = out.length;
    // The group-address path only exists below a group — a slide's own
    // members have nothing above them to address into.
    const cp = prefix.length > 0 ? [...prefix, i] : undefined;
    if ("shape" in child) {
      out.push(...shapeMembers(child.shape, t, cp, context));
    } else if ("picture" in child) {
      const m = pictureMember(child.picture, t, cp, context);
      if (m) out.push(m);
    } else if ("line" in child) {
      out.push(...endpointMembers(child.line, t, cp, {}, context));
    } else if ("connector" in child) {
      out.push(
        ...endpointMembers(
          child.connector,
          t,
          cp,
          geometryOf(child.connector.properties?.geometry),
          context,
        ),
      );
    } else if ("group" in child) {
      out.push(...groupMembers(child.group, t, [...prefix, i], context));
    } else if ("alternateContent" in child) {
      const selected = alternateChildOf(child);
      if (selected) out.push(...childMembers([selected], t, cp ?? [], context));
    } else if ("lockedCanvas" in child) {
      for (const locked of child.lockedCanvas.children ?? []) {
        out.push(
          ...shapeMembers(
            {
              x: locked.x ?? 0,
              y: locked.y ?? 0,
              width: locked.width ?? 0,
              height: locked.height ?? 0,
              properties: locked,
              ...(locked.textBody ? { textBody: { text: locked.textBody } } : {}),
            },
            t,
            cp,
            context,
          ),
        );
      }
    } else if ("smartart" in child) {
      out.push(smartArtMember(child.smartart, t, cp));
    } else if ("video" in child) {
      out.push(mediaMember("video", child.video, t, cp));
    } else if ("audio" in child) {
      out.push(mediaMember("audio", child.audio, t, cp));
    } else if ("table" in child) {
      out.push(tableMember(child.table, t, cp, context));
    } else if ("chart" in child) {
      const c = child.chart;
      out.push({
        kind: "chart",
        x: t.sx * emuOf(c.x) + t.dx,
        y: t.sy * emuOf(c.y) + t.dy,
        width: t.sx * emuOf(c.width),
        height: t.sy * emuOf(c.height),
        ...(cp ? { childPath: cp } : {}),
        chart: resolveChartColors(c, {
          ...context,
          themeColors: chartThemeColorsOf(
            context.themeColors,
            c.themeOverride?.options.colorScheme,
          ),
          colorMapping: c.colorMappingOverride ?? context.colorMapping,
        }),
      });
    } else if ("ole" in child) {
      const o = child.ole;
      const icon = o.iconImage;
      if (!icon) return;
      const member = pictureMember(
        {
          x: o.x,
          y: o.y,
          width: o.width,
          height: o.height,
          data: icon.data,
          type: icon.type as PictureOptions["type"],
          ...(icon.fill ? { fill: icon.fill } : {}),
          ...(icon.outline ? { outline: icon.outline } : {}),
          ...(icon.geometry !== undefined ? { geometry: icon.geometry } : {}),
        },
        t,
        cp,
        context,
      );
      if (member?.kind === "picture" && (member.src || member.line)) out.push(member);
    }
    for (let index = start; index < out.length; index++) {
      out[index]!.sourceChildIndex = i;
    }
  });
  return out;
}

/** One paintable member tagged with the index of the slide child that
 *  projected it; the painter batches and diffs members by that tag. */
export type ProjectedSlideMember = LayoutDrawingMember & { sourceChildIndex?: number };

function groupMembers(
  group: GroupOptions,
  t: Xform,
  path: readonly number[],
  context: TextFieldContext,
): LayoutDrawingMember[] {
  const frame = groupBoxOf(group, t);
  const members = childMembers(group.children ?? [], groupXformOf(group, t), path, context);
  if (!group.rotation) return members;
  return members.map((member) =>
    rotateMemberAbout(
      member,
      { x: frame.x + frame.width / 2, y: frame.y + frame.height / 2 },
      group.rotation!,
    ),
  );
}

function groupXformOf(group: GroupOptions, t: Xform): Xform {
  const frame = groupBoxOf(group, t);
  // An unset chOff/chExt stringifies to the group box itself (identity).
  const chW = emuOf(group.childExtentWidth);
  const chH = emuOf(group.childExtentHeight);
  const sx = chW > 0 ? frame.width / chW : 1;
  const sy = chH > 0 ? frame.height / chH : 1;
  return {
    sx,
    sy,
    dx: frame.x - emuOf(group.childOffsetX) * sx,
    dy: frame.y - emuOf(group.childOffsetY) * sy,
  };
}

/** The group's outer frame before its own spin, in slide-absolute px. */
function groupBoxOf(
  group: GroupOptions,
  t: Xform,
): { x: number; y: number; width: number; height: number } {
  return {
    x: t.sx * emuOf(group.x) + t.dx,
    y: t.sy * emuOf(group.y) + t.dy,
    width: t.sx * emuOf(group.width),
    height: t.sy * emuOf(group.height),
  };
}

/** Fold a group's clockwise spin into a flattened member: its unrotated
 *  extent stays the same, its center rides the group's pivot, and the angle
 *  accumulates so the painter can spin it in one member transform. */
function rotateMemberAbout(
  member: LayoutDrawingMember,
  center: { x: number; y: number },
  delta: number,
): LayoutDrawingMember {
  const cx = member.x + member.width / 2;
  const cy = member.y + member.height / 2;
  const rad = (delta * Math.PI) / 180;
  const dx = cx - center.x;
  const dy = cy - center.y;
  return {
    ...member,
    x: center.x + dx * Math.cos(rad) - dy * Math.sin(rad) - member.width / 2,
    y: center.y + dx * Math.sin(rad) + dy * Math.cos(rad) - member.height / 2,
    rotation: (member.rotation ?? 0) + delta,
  };
}

/** Enter a group's child space: its affine map, plus its own spin ahead of
 *  every outer spin (innermost-first, ready for reverse-order projection). */
function childSpaceOf(group: GroupOptions, parentSpace: MemberSpace): MemberSpace {
  const frame = groupBoxOf(group, parentSpace.xform);
  return {
    xform: groupXformOf(group, parentSpace.xform),
    rotations: group.rotation
      ? [
          {
            x: frame.x + frame.width / 2,
            y: frame.y + frame.height / 2,
            angle: group.rotation,
          },
          ...parentSpace.rotations,
        ]
      : parentSpace.rotations,
  };
}

/** A parent-space point in a group's child space: undo the group's spin,
 *  then its chOff/chExt affine. */
function sourcePointOf(
  group: GroupOptions,
  parentSpace: MemberSpace,
  point: { x: number; y: number },
): { x: number; y: number } {
  const frame = groupBoxOf(group, parentSpace.xform);
  const spun = group.rotation
    ? rotatePoint(
        point,
        { x: frame.x + frame.width / 2, y: frame.y + frame.height / 2 },
        -group.rotation,
      )
    : point;
  const t = groupXformOf(group, parentSpace.xform);
  return { x: (spun.x - t.dx) / t.sx, y: (spun.y - t.dy) / t.sy };
}

/** A final projected point in a leaf's source space: undo every ancestor
 *  spin (outer first), then the composed child affine. */
function sourcePointInSpace(
  space: MemberSpace,
  point: { x: number; y: number },
): { x: number; y: number } {
  let spun = point;
  for (const spin of [...space.rotations].reverse()) spun = rotatePoint(spun, spin, -spin.angle);
  return {
    x: (spun.x - space.xform.dx) / space.xform.sx,
    y: (spun.y - space.xform.dy) / space.xform.sy,
  };
}

/** A slide-space direction in leaf source units. */
function sourceVectorInSpace(
  space: MemberSpace,
  vector: { x: number; y: number },
): { x: number; y: number } {
  let spun = vector;
  for (const spin of [...space.rotations].reverse())
    spun = rotatePoint(spun, { x: 0, y: 0 }, -spin.angle);
  return { x: spun.x / space.xform.sx, y: spun.y / space.xform.sy };
}

function rotatePoint(
  point: { x: number; y: number },
  center: { x: number; y: number },
  angle: number,
): { x: number; y: number } {
  const rad = (angle * Math.PI) / 180;
  const dx = point.x - center.x;
  const dy = point.y - center.y;
  return {
    x: center.x + dx * Math.cos(rad) - dy * Math.sin(rad),
    y: center.y + dx * Math.sin(rad) + dy * Math.cos(rad),
  };
}

/** Move a source-space hit through ancestor spins, outer first; extent and
 *  center semantics match `rotateMemberAbout`. */
function projectHit(hit: MemberHit, space: MemberSpace): MemberHit {
  let rotated = {
    ...hit,
    x: space.xform.sx * hit.x + space.xform.dx,
    y: space.xform.sy * hit.y + space.xform.dy,
    width: space.xform.sx * hit.width,
    height: space.xform.sy * hit.height,
  };
  for (const spin of [...space.rotations].reverse()) {
    const center = { x: rotated.x + rotated.width / 2, y: rotated.y + rotated.height / 2 };
    const point = rotatePoint(center, spin, spin.angle);
    rotated = {
      ...rotated,
      x: point.x - rotated.width / 2,
      y: point.y - rotated.height / 2,
    };
    rotated.groupRotation += spin.angle;
    rotated.rotation += spin.angle;
  }
  return rotated;
}

// ── group member hits ──

/** A group-nested member: the child-index chain under a group child, with
 *  its slide-absolute box and the child-space → px affine map (drag deltas
 *  divide by scale; resize boxes inverse-map through the offset too). */
export interface MemberHit {
  path: number[];
  x: number;
  y: number;
  width: number;
  height: number;
  scale: { sx: number; sy: number };
  offset: { dx: number; dy: number };
  /** The member's own spin plus every ancestor group's folded spin. */
  rotation: number;
  /** Ancestor spin only: screen deltas inverse-rotate by it before they
   *  divide by the group scale. */
  groupRotation: number;
}

/** A group chain's leaf space: the affine map plus group spins in
 *  innermost-first order. Hit boxes live before those spins; projected
 *  members live after them. */
interface MemberSpace {
  xform: Xform;
  rotations: { x: number; y: number; angle: number }[];
}

const contains = (
  b: { x: number; y: number; width: number; height: number },
  x: number,
  y: number,
): boolean => x >= b.x && x <= b.x + b.width && y >= b.y && y <= b.y + b.height;

/** A point in an unrotated box after undoing the member's own spin. Group
 *  ancestor spins are already removed while mapping into child space. */
function containsRotated(
  box: { x: number; y: number; width: number; height: number },
  rotation: number,
  x: number,
  y: number,
): boolean {
  if (!rotation) return contains(box, x, y);
  const center = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  const rad = (-rotation * Math.PI) / 180;
  const dx = x - center.x;
  const dy = y - center.y;
  return contains(
    box,
    center.x + dx * Math.cos(rad) - dy * Math.sin(rad),
    center.y + dx * Math.sin(rad) + dy * Math.cos(rad),
  );
}

/** The child's unspun outer box in its own parent's child coordinates —
 *  transform variants from their xfrm, lines and connectors from endpoints.
 *  Unplaced children hit nothing. */
function childBoxUnder(
  child: SlideChild,
): { x: number; y: number; width: number; height: number } | null {
  if ("alternateContent" in child) {
    const selected = alternateChildOf(child);
    return selected ? childBoxUnder(selected) : null;
  }
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    const x1 = emuOf(o.x1);
    const y1 = emuOf(o.y1);
    const x2 = emuOf(o.x2);
    const y2 = emuOf(o.y2);
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
              : "smartart" in child
                ? child.smartart
                : "video" in child
                  ? child.video
                  : "audio" in child
                    ? child.audio
                    : "ole" in child
                      ? child.ole
                      : "lockedCanvas" in child
                        ? child.lockedCanvas
                        : undefined;
  if (!g) return null;
  return {
    x: emuOf(g.x),
    y: emuOf(g.y),
    width: emuOf(g.width),
    height: emuOf(g.height),
  };
}

const hitOf = (
  path: number[],
  box: { x: number; y: number; width: number; height: number },
  t: Xform,
  rotation: number,
): MemberHit => ({
  path,
  ...box,
  scale: { sx: t.sx, sy: t.sy },
  offset: { dx: t.dx, dy: t.dy },
  rotation,
  groupRotation: 0,
});

/** The leaf's own a:xfrm spin (lines and graphic frames have none). */
function ownRotationOf(child: SlideChild): number {
  if ("shape" in child) return child.shape.rotation ?? 0;
  if ("picture" in child) return child.picture.rotation ?? 0;
  if ("group" in child) return child.group.rotation ?? 0;
  return 0;
}

/** The deepest group-nested member whose box contains the slide-local point
 *  — a child-index chain under the group child — or null: the point misses
 *  every member and the group itself is the hit. A nested group whose box
 *  contains the point but no member of its own does reads as the group. */
export function memberAt(child: SlideChild, x: number, y: number): MemberHit | null {
  if (!("group" in child)) return null;
  return hitGroupAt(child.group, { xform: IDENTITY, rotations: [] }, x, y, []);
}

function hitMemberIn(
  children: readonly SlideChild[],
  space: MemberSpace,
  x: number,
  y: number,
  prefix: number[],
): MemberHit | null {
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i]!;
    if (hiddenOf(c)) continue;
    const box = childBoxUnder(c);
    if (!box || !containsRotated(box, ownRotationOf(c), x, y)) continue;
    if ("group" in c) {
      const deeper = hitGroupAt(c.group, space, x, y, [...prefix, i]);
      if (deeper) return deeper;
    }
    return hitOf([...prefix, i], box, space.xform, ownRotationOf(c));
  }
  return null;
}

/** Resolve a hit in a group's unspun child space, then fold this group's
 *  spin and every outer ancestor's spin back onto it. */
function hitGroupAt(
  group: GroupOptions,
  parentSpace: MemberSpace,
  x: number,
  y: number,
  prefix: number[],
): MemberHit | null {
  const space = childSpaceOf(group, parentSpace);
  const local = sourcePointOf(group, parentSpace, { x, y });
  const hit = hitMemberIn(group.children ?? [], space, local.x, local.y, prefix);
  return hit && projectHit(hit, space);
}

/** The member at `path` (child indexes under the group child) — its box and
 *  scale, or null when the path dangles (the deck changed under the
 *  selection). */
export function memberByPath(child: SlideChild, path: readonly number[]): MemberHit | null {
  if (!("group" in child) || path.length === 0) return null;
  const found = memberSpaceByPath(child, path);
  if (!found) return null;
  const box = childBoxUnder(found.node);
  if (!box) return null;
  return projectHit(
    hitOf([...path], box, found.space.xform, ownRotationOf(found.node)),
    found.space,
  );
}

function memberSpaceByPath(
  root: SlideChild & { group: GroupOptions },
  path: readonly number[],
): { node: SlideChild; space: MemberSpace } | null {
  let group = root.group;
  let parentSpace: MemberSpace = { xform: IDENTITY, rotations: [] };
  for (let depth = 0; depth < path.length; depth++) {
    const space = childSpaceOf(group, parentSpace);
    const node = group.children?.[path[depth]!];
    if (!node) return null;
    if (depth === path.length - 1) return { node, space };
    if (!("group" in node)) return null;
    parentSpace = space;
    group = node.group;
  }
  return null;
}

/** A px measure back to the integer EMU the stringify side emits. */
const emu = (px: number): number => Math.round(px * EMU_PER_PX);

/** Resize a group member found at `path` to the slide-absolute `box`.
 *  Transform children inverse-map their own box; lines inverse-map each
 *  endpoint while preserving which endpoint owns which box corner. */
export function resizeMemberByPath(
  child: SlideChild,
  path: readonly number[],
  box: { x: number; y: number; width: number; height: number },
): void {
  if (!("group" in child) || path.length === 0) return;
  const found = memberSpaceByPath(child, path);
  if (!found) return;
  const t = found.space.xform;
  const center = sourcePointInSpace(found.space, {
    x: box.x + box.width / 2,
    y: box.y + box.height / 2,
  });
  const sourceWidth = box.width / t.sx;
  const sourceHeight = box.height / t.sy;
  if ("line" in found.node || "connector" in found.node) {
    const o = "line" in found.node ? found.node.line : found.node.connector;
    const spin = memberByPath(child, path)?.rotation ?? 0;
    const sx = emuOf(o.x1) <= emuOf(o.x2) ? 0 : 1;
    const sy = emuOf(o.y1) <= emuOf(o.y2) ? 0 : 1;
    const endpoint = (cornerX: number, cornerY: number) => {
      const stored = rotatePoint(
        { x: cornerX, y: cornerY },
        { x: box.x + box.width / 2, y: box.y + box.height / 2 },
        -spin,
      );
      return sourcePointInSpace(found.space, stored);
    };
    const a = endpoint(box.x + sx * box.width, box.y + sy * box.height);
    const b = endpoint(box.x + (1 - sx) * box.width, box.y + (1 - sy) * box.height);
    o.x1 = emu(a.x);
    o.y1 = emu(a.y);
    o.x2 = emu(b.x);
    o.y2 = emu(b.y);
    return;
  }
  const g =
    "shape" in found.node
      ? found.node.shape
      : "picture" in found.node
        ? found.node.picture
        : "group" in found.node
          ? found.node.group
          : "table" in found.node
            ? found.node.table
            : "chart" in found.node
              ? found.node.chart
              : undefined;
  if (!g) return;
  g.x = emu(center.x - sourceWidth / 2);
  g.y = emu(center.y - sourceHeight / 2);
  g.width = emu(sourceWidth);
  g.height = emu(sourceHeight);
}

/** Move a group member by a slide-absolute screen delta. */
export function offsetMemberByPath(
  child: SlideChild,
  path: readonly number[],
  dx: number,
  dy: number,
): void {
  if (!("group" in child) || path.length === 0) return;
  const found = memberSpaceByPath(child, path);
  if (!found) return;
  const local = sourceVectorInSpace(found.space, { x: dx, y: dy });
  const node = found.node;
  if ("line" in node || "connector" in node) {
    const o = "line" in node ? node.line : node.connector;
    o.x1 = emu(emuOf(o.x1) + local.x);
    o.y1 = emu(emuOf(o.y1) + local.y);
    o.x2 = emu(emuOf(o.x2) + local.x);
    o.y2 = emu(emuOf(o.y2) + local.y);
    return;
  }
  const g =
    "shape" in node
      ? node.shape
      : "picture" in node
        ? node.picture
        : "group" in node
          ? node.group
          : "table" in node
            ? node.table
            : "chart" in node
              ? node.chart
              : undefined;
  if (!g) return;
  g.x = emu(emuOf(g.x) + local.x);
  g.y = emu(emuOf(g.y) + local.y);
}
