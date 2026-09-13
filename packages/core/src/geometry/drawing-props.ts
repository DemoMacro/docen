// Format-neutral DrawingML paint extraction: the spPr JSON shape office-open
// parses (fill / a:ln / a:effectLst, EMU measures) → the flat props the
// drawing members carry. Duck-typed on purpose — one reader serves docx
// (wps:spPr), pptx (p:spPr) and xlsx shapes without importing their option
// types. Lives on the Node-safe geometry subpath: the layout projections run
// headless.

import { emuToPx, type LayoutDrawingMember, type LayoutDrawingShadow } from "@docen/layout";

// ── loose-shape guards ──

export type Rec = Record<string, unknown>;

/** The options unions are structurally loose at their edges (optional
 *  everything, per-side sub-objects); this guard narrows an unknown pick to a
 *  record so the readers below work without per-site casts. */
export function isRecord(v: unknown): v is Rec {
  return !!v && typeof v === "object";
}

export const num = (v: unknown): number | undefined => (typeof v === "number" ? v : undefined);

export const str = (v: unknown): string | undefined => (typeof v === "string" && v ? v : undefined);

// ── universal-measure parsing (number = native unit, string = UM) ──

const UM_IN_TWIPS = { pt: 20, pc: 240, in: 1440, mm: 1440 / 25.4, cm: 1440 / 2.54, px: 15 };
const UM_RE = /^(-?[\d.]+)(pt|pc|in|mm|cm|px)$/;

function measureTwip(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const m = UM_RE.exec(v);
  return m ? Number(m[1]) * UM_IN_TWIPS[m[2] as keyof typeof UM_IN_TWIPS] : undefined;
}

/** A measure field whose native unit is EMU (drawing extents): number passes,
 *  UM resolves (px at 96 dpi). */
export function measureEmu(v: unknown): number | undefined {
  const n = num(v);
  if (n != null) return n;
  if (typeof v !== "string") return undefined;
  const tw = measureTwip(v);
  return tw != null ? (tw / 1440) * 914400 : undefined;
}

/** A color field: bare hex string, or the round-trip object shape
 *  (`{value}` on outline/fill colors, `{val, themeColor}` on run colors —
 *  the parse emits both key spellings). Theme-only colors resolve later. */
export function colorOf(v: unknown): string | undefined {
  if (typeof v === "string") return v === "auto" ? undefined : v;
  if (isRecord(v)) return str(v.value) ?? str(v.val);
  return undefined;
}

// ── spPr paint extraction ──

/** Solid fill (FillOptions union) → hex; every other variant (none/gradient/
 *  picture) carries no paintable flat color → undefined. The bare-string
 *  shorthand is a solid color by convention. */
export function solidFillOf(fill: unknown): string | undefined {
  if (typeof fill === "string") return colorOf(fill);
  return isRecord(fill) && fill.type === "solid" ? colorOf(fill.color) : undefined;
}

/** Solid-fill opacity from the color's alpha transform (integer percent,
 *  0-100). The flat-color painter can only fade a whole fill, so modulate/
 *  offset stacks collapse to the base alpha; anything else passes opaque. */
export function fillOpacityOf(fill: unknown): number | undefined {
  if (!isRecord(fill) || fill.type !== "solid") return undefined;
  const c = fill.color;
  const a = isRecord(c) && isRecord(c.transforms) ? num(c.transforms.alpha) : undefined;
  return a != null && a < 100 ? Math.max(0, Math.min(1, a / 100)) : undefined;
}

/** The gradient stop closest to the middle position — the flattest honest
 *  color for a gradient the painter cannot stroke. */
function midStopOf(gradient: unknown): string | undefined {
  const stops = isRecord(gradient) && Array.isArray(gradient.stops) ? gradient.stops : undefined;
  if (!stops) return undefined;
  let best: { pos: number; color: string } | undefined;
  for (const stop of stops) {
    if (!isRecord(stop)) continue;
    const pos = num(stop.position);
    const color = colorOf(stop.color);
    if (pos == null || color == null) continue;
    if (!best || Math.abs(pos - 50) < Math.abs(best.pos - 50)) best = { pos, color };
  }
  return best?.color;
}

/** A line end (a:headEnd/a:tailEnd) → the arrow type plus its resolved length
 *  in px: DrawingML sizes the arrow in stroke widths (sm/med/lg length over
 *  sm/med/lg width — POI's ArrowScale), med/med ≈ 3× the stroke. "none" (the
 *  schema default) stays undefined. */
function lineEndOf(end: unknown, linePx: number): { type: string; px: number } | undefined {
  if (!isRecord(end)) return undefined;
  const type = str(end.type);
  if (!type || type === "none") return undefined;
  const len = end.length === "small" ? 2 : end.length === "large" ? 4.5 : 3;
  const wide = end.width === "small" ? 0.8 : end.width === "large" ? 1.25 : 1;
  return { type, px: linePx * len * wide };
}

/** Outline stroke (a:ln): px width + color + the line-dressing tokens the
 *  painter maps (cap/join full-word, dash the OOXML prstDash token). A
 *  gradient stroke flattens to its middle stop's color — the painter strokes
 *  flat colors, and a line's gradient averages visually to its middle.
 *  Shared by shape members and picture borders. */
export function outlineOf(outline: unknown):
  | {
      px: number;
      color?: string;
      cap?: "round" | "square" | "flat";
      join?: "round" | "bevel" | "miter";
      dash?: string;
      headEnd?: { type: string; px: number };
      tailEnd?: { type: string; px: number };
    }
  | undefined {
  if (!isRecord(outline) || outline.type === "noFill") return undefined;
  const widthEmu = num(outline.width);
  if (widthEmu == null) return undefined;
  const cap =
    outline.cap === "round" || outline.cap === "square" || outline.cap === "flat"
      ? outline.cap
      : undefined;
  const join =
    outline.join === "round" || outline.join === "bevel" || outline.join === "miter"
      ? outline.join
      : undefined;
  return {
    px: emuToPx(widthEmu),
    color: colorOf(outline.color) ?? midStopOf(outline.gradientFill),
    cap,
    join,
    dash: str(outline.dash),
    headEnd: lineEndOf(outline.headEnd, emuToPx(widthEmu)),
    tailEnd: lineEndOf(outline.tailEnd, emuToPx(widthEmu)),
  };
}

/** A straight-line preset's segment: the box diagonal, corner to corner
 *  (top-left first). The xfrm flips are NOT baked in here — the painter's
 *  mirror group is the single flip consumer (baking them in too would mirror
 *  twice and cancel the flip out). */
export function linePathData(width: number, height: number): string {
  const p = (v: number): string => String(Math.round(v * 100) / 100);
  return `M 0 0 L ${p(width)} ${p(height)}`;
}

/** One line-end arrow expanded into its own paintable path (box-local px).
 *  triangle/stealth/diamond/oval fill the line color; "arrow" is the open
 *  chevron, stroked. Returns undefined for types without geometry. */
function lineEndMember(
  end: { type: string; px: number },
  ex: number,
  ey: number,
  ux: number,
  uy: number,
  color: string | undefined,
  linePx: number,
): LayoutDrawingMember | undefined {
  const L = end.px;
  const W = L * 0.42;
  const nx = -uy;
  const ny = ux;
  const stroke = { px: Math.max(1, linePx), ...(color ? { color } : {}) };
  // Polygon vertices (box-local): tip plus the trailing points per shape.
  let pts: [number, number][];
  let closed = true;
  switch (end.type) {
    case "triangle":
      pts = [
        [ex, ey],
        [ex - ux * L + nx * W, ey - uy * L + ny * W],
        [ex - ux * L - nx * W, ey - uy * L - ny * W],
      ];
      break;
    case "stealth":
      pts = [
        [ex, ey],
        [ex - ux * L + nx * W, ey - uy * L + ny * W],
        [ex - ux * L * 0.5, ey - uy * L * 0.5],
        [ex - ux * L - nx * W, ey - uy * L - ny * W],
      ];
      break;
    case "diamond":
      pts = [
        [ex, ey],
        [ex - ux * L * 0.5 + nx * W, ey - uy * L * 0.5 + ny * W],
        [ex - ux * L, ey - uy * L],
        [ex - ux * L * 0.5 - nx * W, ey - uy * L * 0.5 - ny * W],
      ];
      break;
    case "oval":
      pts = [[ex - ux * L * 0.5, ey - uy * L * 0.5]];
      break;
    case "arrow":
      // The open chevron: two strokes meeting at the tip.
      pts = [
        [ex - ux * L + nx * W, ey - uy * L + ny * W],
        [ex, ey],
        [ex - ux * L - nx * W, ey - uy * L - ny * W],
      ];
      closed = false;
      break;
    default:
      return undefined;
  }
  const box = (() => {
    const xs = pts.map((q) => q[0]);
    const ys = pts.map((q) => q[1]);
    if (end.type === "oval") {
      return {
        x: xs[0] - W,
        y: ys[0] - W,
        width: W * 2,
        height: W * 2,
      };
    }
    return {
      x: Math.min(...xs),
      y: Math.min(...ys),
      width: Math.max(...xs) - Math.min(...xs),
      height: Math.max(...ys) - Math.min(...ys),
    };
  })();
  const p = (v: number): string => String(Math.round(v * 100) / 100);
  const at = (q: [number, number]): string => `${p(q[0] - box.x)} ${p(q[1] - box.y)}`;
  let d: string;
  if (end.type === "oval") {
    // A circle of radius W — four cubic segments (κ = 0.5523).
    const [cx, cy] = pts[0];
    const k = W * 0.5523;
    const rx = cx - box.x;
    const ry = cy - box.y;
    d =
      `M ${p(rx + W)} ${p(ry)}` +
      ` C ${p(rx + W)} ${p(ry + k)} ${p(rx + k)} ${p(ry + W)} ${p(rx)} ${p(ry + W)}` +
      ` C ${p(rx - k)} ${p(ry + W)} ${p(rx - W)} ${p(ry + k)} ${p(rx - W)} ${p(ry)}` +
      ` C ${p(rx - W)} ${p(ry - k)} ${p(rx - k)} ${p(ry - W)} ${p(rx)} ${p(ry - W)}` +
      ` C ${p(rx + k)} ${p(ry - W)} ${p(rx + W)} ${p(ry - k)} ${p(rx + W)} ${p(ry)} Z`;
  } else {
    d = `M ${at(pts[0])}${pts
      .slice(1)
      .map((q) => ` L ${at(q)}`)
      .join("")}${closed ? " Z" : ""}`;
  }
  return {
    kind: "path",
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    d,
    ...(end.type === "arrow" ? (color ? { line: stroke } : {}) : color ? { fill: color } : {}),
  };
}

/** The straight-line member's line-end arrows: expanded into their own fill
 *  members at both ends of the segment (a:headEnd/a:tailEnd). */
export function lineEndMembersOf(
  line:
    | {
        color?: string;
        px: number;
        headEnd?: { type: string; px: number };
        tailEnd?: { type: string; px: number };
      }
    | undefined,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): LayoutDrawingMember[] {
  if (!line) return [];
  const len = Math.hypot(x1 - x0, y1 - y0);
  if (len < 1e-6) return [];
  const ux = (x1 - x0) / len;
  const uy = (y1 - y0) / len;
  const head = line.headEnd
    ? lineEndMember(line.headEnd, x0, y0, -ux, -uy, line.color, line.px)
    : undefined;
  const tail = line.tailEnd
    ? lineEndMember(line.tailEnd, x1, y1, ux, uy, line.color, line.px)
    : undefined;
  return [...(head ? [head] : []), ...(tail ? [tail] : [])];
}

/** a:outerShdw → the flat shadow: dist+dir resolve into the px offset
 *  (DrawingML measures clockwise from the 3-o'clock direction, and screen y
 *  grows downward, so x = dist·cos, y = dist·sin directly). */
function shadowOf(v: Rec): LayoutDrawingShadow | undefined {
  const dist = num(v.distance);
  const dir = num(v.direction) ?? 0;
  const blur = num(v.blurRadius);
  const x = dist != null ? Math.cos((dir * Math.PI) / 180) * emuToPx(dist) : 0;
  const y = dist != null ? Math.sin((dir * Math.PI) / 180) * emuToPx(dist) : 0;
  const color = colorOf(v.color);
  const alpha = isRecord(v.color) ? num((v.color.transforms as Rec | undefined)?.alpha) : undefined;
  if (!x && !y && !blur && !color) return undefined;
  return {
    x,
    y,
    blur: blur != null ? emuToPx(blur) : 0,
    ...(color ? { color } : {}),
    ...(alpha != null && alpha < 100 ? { opacity: Math.max(0, alpha / 100) } : {}),
  };
}

/** An effects list's outer shadow (a:effectLst a:outerShdw). */
export function outerShadowOf(effects: unknown): LayoutDrawingShadow | undefined {
  const shdw = isRecord(effects) ? effects.outerShadow : undefined;
  return isRecord(shdw) ? shadowOf(shdw) : undefined;
}
