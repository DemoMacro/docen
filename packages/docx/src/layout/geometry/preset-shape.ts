import type { GeometryGuide } from "@office-open/core";

import { PRESET_SHAPE_DEFS, type PresetShapePathDef } from "./preset-shape-data";

/**
 * Preset shape geometry evaluator — expands a drawingml preset geometry
 * (a:prstGeom token + optional w:avLst overrides) into SVG path data, the same
 * contract the path members already paint. Formula and path-op semantics follow
 * ECMA-376 §20.1.9 as implemented by Apache POI's sl.draw.geom: angles are in
 * 1/60000 degree, sin/cos/tan multiply by their first operand, at2 returns
 * 1/60000 degree, cat2/sat2 are cos/sin of atan2, and a path with w/h
 * attributes evaluates every guide (including l/t/r/b) in that coordinate
 * space before the result is stretched to the box.
 */

/** OOXML angle unit: 1/60000 degree. */
const DEG = 60000;
const QUARTER = 90 * DEG;

/** One paintable outline of a preset shape, in box coordinates (0,0 … width,height). */
export interface PresetShapeOutline {
  /** SVG path data (`M 0 100 L … Z`), already scaled to the box. */
  readonly d: string;
  /** Paint with the shape's fill color. */
  readonly fill: boolean;
  /** Paint with the shape's outline color. */
  readonly stroke: boolean;
}

interface Guides {
  get(name: string): number | undefined;
  set(name: string, value: number): void;
}

// ECMA-376 built-in guide values, evaluated against the current frame (the
// box, or the path's coordinate space when the path carries w/h).
const builtIns = (w: number, h: number): Map<string, number> => {
  const ss = Math.min(w, h);
  return new Map<string, number>([
    ["l", 0],
    ["t", 0],
    ["r", w],
    ["b", h],
    ["w", w],
    ["h", h],
    ["hc", w / 2],
    ["vc", h / 2],
    ["ls", Math.max(w, h)],
    ["ss", ss],
    ...[2, 3, 4, 5, 6, 8, 10, 32].map((d) => [`wd${d}`, w / d] as const),
    ...[2, 3, 4, 5, 6, 8].map((d) => [`hd${d}`, h / d] as const),
    ...[2, 4, 6, 8, 16, 32].map((d) => [`ssd${d}`, ss / d] as const),
    ["cd2", 180 * DEG],
    ["cd4", QUARTER],
    ["cd8", 45 * DEG],
    ["3cd4", 270 * DEG],
    ["3cd8", 135 * DEG],
    ["5cd8", 225 * DEG],
    ["7cd8", 315 * DEG],
  ]);
};

// `*/ x y z` → x*y/z, `+-` → x+y-z, `+/` → (x+y)/z, `?:` → x>0?y:z; angles in
// 1/60000 degree (POI GuideIf semantics).
const formula = (op: string, a: number, b: number, c: number): number => {
  switch (op) {
    case "*/":
      return c === 0 ? 0 : (a * b) / c;
    case "+-":
      return a + b - c;
    case "+/":
      return c === 0 ? 0 : (a + b) / c;
    case "?:":
      return a > 0 ? b : c;
    case "abs":
      return Math.abs(a);
    case "at2":
      return ((Math.atan2(b, a) * 180) / Math.PI) * DEG;
    case "cat2":
      return a * Math.cos(Math.atan2(c, b));
    case "sat2":
      return a * Math.sin(Math.atan2(c, b));
    case "cos":
      return a * Math.cos((b / DEG) * (Math.PI / 180));
    case "sin":
      return a * Math.sin((b / DEG) * (Math.PI / 180));
    case "tan":
      return a * Math.tan((b / DEG) * (Math.PI / 180));
    case "max":
      return Math.max(a, b);
    case "min":
      return Math.min(a, b);
    case "mod":
      return Math.sqrt(a * a + b * b + c * c);
    case "pin":
      return Math.max(a, Math.min(b, c));
    case "sqrt":
      return Math.sqrt(a);
    default:
      return a;
  }
};

const value = (token: string, guides: Guides): number => {
  const literal = Number(token);
  if (!Number.isNaN(literal)) return literal;
  return guides.get(token) ?? 0;
};

// "name op operand…" — evaluates against the current guides and binds the name.
const guide = (entry: string, guides: Guides): void => {
  const tokens = entry.split(" ");
  const op = tokens[1] ?? "val";
  const a = value(tokens[2] ?? "0", guides);
  guides.set(
    tokens[0],
    formula(op, a, value(tokens[3] ?? "0", guides), value(tokens[4] ?? "0", guides)),
  );
};

const r2 = (v: number): string => String(Math.round(v * 100) / 100);

// One path's token stream → SVG `d`. Coordinates are evaluated in the path
// space (frame) and stretched to the box; arcTo derives its center from the
// current point (it always starts on the ellipse at stAng) and emits ≤90°
// cubic segments, which the SVG renderer consumes without arc support.
const pathData = (
  path: PresetShapePathDef,
  guides: Guides,
  width: number,
  height: number,
): string => {
  const fw = path.w || width;
  const fh = path.h || height;
  const sx = width / fw;
  const sy = height / fh;
  const tokens = path.d.split(" ");
  let d = "";
  let x = 0;
  let y = 0;
  const xy = (px: number, py: number): void => {
    x = px;
    y = py;
    d += ` ${r2(x * sx)} ${r2(y * sy)}`;
  };
  for (let i = 0; i < tokens.length;) {
    const op = tokens[i++];
    switch (op) {
      case "x":
        d += " Z";
        break;
      case "m":
      case "l": {
        d += op === "m" ? " M" : " L";
        xy(value(tokens[i++], guides), value(tokens[i++], guides));
        break;
      }
      case "c":
      case "q": {
        d += op === "c" ? " C" : " Q";
        for (let p = 0; p < (op === "c" ? 3 : 2); p++) {
          xy(value(tokens[i++], guides), value(tokens[i++], guides));
        }
        break;
      }
      case "a": {
        const rx = value(tokens[i++], guides);
        const ry = value(tokens[i++], guides);
        const st = (value(tokens[i++], guides) / DEG) * (Math.PI / 180);
        const sw = (value(tokens[i++], guides) / DEG) * (Math.PI / 180);
        // The current point is on the ellipse at stAng — the center follows.
        const cx = x - rx * Math.cos(st);
        const cy = y - ry * Math.sin(st);
        const segments = Math.ceil(Math.abs(sw) / (Math.PI / 2));
        const step = segments ? sw / segments : 0;
        const k = (4 / 3) * Math.tan(step / 4);
        for (let s = 1; s <= segments; s++) {
          const t0 = st + step * (s - 1);
          const t1 = st + step * s;
          const x1 = cx + rx * Math.cos(t0);
          const y1 = cy + ry * Math.sin(t0);
          x = cx + rx * Math.cos(t1);
          y = cy + ry * Math.sin(t1);
          d += " C";
          d += ` ${r2((x1 + k * -rx * Math.sin(t0)) * sx)} ${r2((y1 + k * ry * Math.cos(t0)) * sy)}`;
          d += ` ${r2((x - k * -rx * Math.sin(t1)) * sx)} ${r2((y - k * ry * Math.cos(t1)) * sy)}`;
          xy(x, y);
        }
        break;
      }
      default:
        break;
    }
  }
  return d.trim();
};

/**
 * Expand a preset geometry into paintable outlines.
 *
 * @param preset - `a:prstGeom` token (e.g. "star5"); unknown tokens return undefined.
 * @param width - Box width in the caller's unit (any unit — coordinates keep the scale).
 * @param height - Box height in the same unit as width.
 * @param adjustmentValues - Document avLst entries; a name replaces the preset default of the same name.
 * @returns One outline per ECMA-376 path with fill "norm"/"none", in document order. The
 *   darken/darkenLess/lighten/lightenLess shading paths are not returned (they stay in the data module).
 */
export function presetShapePaths(
  preset: string,
  width: number,
  height: number,
  adjustmentValues?: readonly GeometryGuide[],
): PresetShapeOutline[] | undefined {
  const def = PRESET_SHAPE_DEFS[preset];
  if (!def) return undefined;
  const guides = builtIns(width, height);
  for (const entry of def.av) guide(entry, guides);
  for (const { name, formula: fmla } of adjustmentValues ?? []) {
    if (!name) continue;
    const tokens = fmla.split(" ");
    guides.set(
      name,
      formula(
        tokens[0],
        value(tokens[1] ?? "0", guides),
        value(tokens[2] ?? "0", guides),
        value(tokens[3] ?? "0", guides),
      ),
    );
  }
  for (const entry of def.gd) guide(entry, guides);
  return def.paths
    .filter((p) => p.fill === "norm" || p.fill === "none")
    .map((p) => ({
      d: pathData(p, guides, width, height),
      fill: p.fill === "norm",
      stroke: p.stroke,
    }));
}
