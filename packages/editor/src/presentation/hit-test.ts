// Slide-object hit-testing and geometry write-back over the source slide
// JSON. The projected members are paint-shaped (and an arrowed line emits
// several members per child), so addressing goes through the slide's own
// children instead: each paintable child contributes one hit box in
// slide-absolute px, and edits write the child's geometry fields back in
// EMU. Paintable = what the projection draws today (shape/picture/line/
// connector/group/table/chart); unpainted variants stay unselectable until
// their painter lands.

import { measureEmu } from "@docen/core/geometry";
import { EMU_PER_PX, emuToPx } from "@docen/layout";
import type { SlideChild, SlideOptions } from "@docen/pptx";

import type { Box } from "../drawing/geometry";

/** One selectable top-level object: its slide-absolute box (plus spin) and
 *  the index into the slide's source children. */
export interface SlideHit {
  child: number;
  box: Box;
  /** Clockwise degrees about the box center (the overlay's rotate handle
   *  shows this angle); endpoint-model children never spin. */
  rotation: number;
}

/** A measure field to px (the projection's own conversion). */
const px = (v: unknown): number => emuToPx(measureEmu(v) ?? 0);

/** A px delta/size back to the integer EMU the stringify side emits. */
const emu = (v: number): number => Math.round(v * EMU_PER_PX);

/** The hit boxes for a slide's paintable children, in source order. */
export function slideHits(slide: SlideOptions): SlideHit[] {
  const hits: SlideHit[] = [];
  (slide.children ?? []).forEach((child, i) => {
    const box = childBox(child);
    if (box) hits.push({ child: i, box, rotation: rotationOf(child) });
  });
  return hits;
}

/** The topmost child whose box contains the slide-local point, or -1 —
 *  later children paint on top, so the scan runs tail-first. */
export function hitSlide(hits: readonly SlideHit[], x: number, y: number): number {
  for (let i = hits.length - 1; i >= 0; i--) {
    const b = hits[i]!.box;
    if (x >= b.x && x <= b.x + b.width && y >= b.y && y <= b.y + b.height) return hits[i]!.child;
  }
  return -1;
}

function childBox(child: SlideChild): Box | null {
  // Lines and connectors address endpoints, not a transform: the box is
  // their bounding rectangle.
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    const x1 = px(o.x1);
    const y1 = px(o.y1);
    const x2 = px(o.x2);
    const y2 = px(o.y2);
    return {
      x: Math.min(x1, x2),
      y: Math.min(y1, y2),
      width: Math.abs(x2 - x1),
      height: Math.abs(y2 - y1),
    };
  }
  const t = transformOf(child);
  if (!t) return null;
  return { x: px(t.x), y: px(t.y), width: px(t.width), height: px(t.height) };
}

// ── write-back ──

/** Move a child by a px delta. Values normalize to EMU numbers — a source
 *  that parsed into a measure string re-lands as the integer EMU. */
export function offsetChild(child: SlideChild, dx: number, dy: number): void {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    o.x1 = emu(px(o.x1) + dx);
    o.y1 = emu(px(o.y1) + dy);
    o.x2 = emu(px(o.x2) + dx);
    o.y2 = emu(px(o.y2) + dy);
    return;
  }
  const t = transformOf(child);
  if (!t) return;
  t.x = emu(px(t.x) + dx);
  t.y = emu(px(t.y) + dy);
}

/** Resize a child to the dragged box. Lines keep their corner direction:
 *  each endpoint re-lands on the matching corner of the new box. */
export function resizeChild(child: SlideChild, box: Box): void {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    const sx = px(o.x1) <= px(o.x2) ? 0 : 1;
    const sy = px(o.y1) <= px(o.y2) ? 0 : 1;
    o.x1 = emu(box.x + sx * box.width);
    o.y1 = emu(box.y + sy * box.height);
    o.x2 = emu(box.x + (1 - sx) * box.width);
    o.y2 = emu(box.y + (1 - sy) * box.height);
    return;
  }
  const t = transformOf(child);
  if (!t) return;
  t.x = emu(box.x);
  t.y = emu(box.y);
  t.width = emu(box.width);
  t.height = emu(box.height);
}

function transformOf(
  child: SlideChild,
): { x?: unknown; y?: unknown; width?: unknown; height?: unknown; rotation?: number } | undefined {
  if ("shape" in child) return child.shape;
  if ("picture" in child) return child.picture;
  if ("group" in child) return child.group;
  if ("table" in child) return child.table;
  if ("chart" in child) return child.chart;
  return undefined;
}

/** Whether the child carries a rotation field (a:xfrm @rot) — table and
 *  chart frames have none, so the rotate gesture must not write one. */
function rotatable(child: SlideChild): boolean {
  return "shape" in child || "picture" in child || "group" in child;
}

/** The child's own spin in degrees (transform children only). */
function rotationOf(child: SlideChild): number {
  return transformOf(child)?.rotation ?? 0;
}

/** Spin a child's box about its center by a degree delta. Transform
 *  children accumulate the spin; lines rotate their endpoints around the
 *  bounding-box center (the endpoint model has no angle of its own). */
export function rotateChild(child: SlideChild, delta: number): void {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    const cx = (px(o.x1) + px(o.x2)) / 2;
    const cy = (px(o.y1) + px(o.y2)) / 2;
    const rad = (delta * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const spin = (x: number, y: number): { x: number; y: number } => ({
      x: cx + (x - cx) * cos - (y - cy) * sin,
      y: cy + (x - cx) * sin + (y - cy) * cos,
    });
    const a = spin(px(o.x1), px(o.y1));
    const b = spin(px(o.x2), px(o.y2));
    o.x1 = emu(a.x);
    o.y1 = emu(a.y);
    o.x2 = emu(b.x);
    o.y2 = emu(b.y);
    return;
  }
  const t = transformOf(child);
  if (!t || !rotatable(child)) return;
  t.rotation = Math.round(((t.rotation ?? 0) + delta) * 100) / 100;
}

// ── undo snapshots ──

/** The child's addressable geometry fields, verbatim (EMU numbers or
 *  measure strings — restore writes them back untouched). */
export interface GeometrySnapshot {
  x?: unknown;
  y?: unknown;
  width?: unknown;
  height?: unknown;
  x1?: unknown;
  y1?: unknown;
  x2?: unknown;
  y2?: unknown;
  rotation?: number;
}

/** Copy the geometry fields an edit may touch. */
export function captureGeometry(child: SlideChild): GeometrySnapshot {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    return { x1: o.x1, y1: o.y1, x2: o.x2, y2: o.y2 };
  }
  const t = transformOf(child);
  return t ? { x: t.x, y: t.y, width: t.width, height: t.height, rotation: t.rotation } : {};
}

/** Write a snapshot's fields back (the undo/redo leg). Snapshots hold the
 *  fields verbatim, so each write casts back to the field's own type. */
export function restoreGeometry(child: SlideChild, snap: GeometrySnapshot): void {
  if ("line" in child || "connector" in child) {
    const o = "line" in child ? child.line : child.connector;
    o.x1 = snap.x1 as typeof o.x1;
    o.y1 = snap.y1 as typeof o.y1;
    o.x2 = snap.x2 as typeof o.x2;
    o.y2 = snap.y2 as typeof o.y2;
    return;
  }
  const t = transformOf(child);
  if (!t) return;
  t.x = snap.x as typeof t.x;
  t.y = snap.y as typeof t.y;
  t.width = snap.width as typeof t.width;
  t.height = snap.height as typeof t.height;
  t.rotation = snap.rotation;
}
