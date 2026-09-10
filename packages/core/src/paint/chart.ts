import type { LayoutDrawingMember } from "@docen/layout";
import { Ellipse, Group, Path as LeaferPath, Rect, Text, type IGroup } from "leafer-ui";

import type { ChartHitContext, ChartPartHit, ChartPartShape } from "./context";

// ── chart member painter ──
//
// Draws a `kind: "chart"` member from its verbatim ChartSpaceOptions payload
// (the office-open core chart model, consumed structurally — this package
// holds no office-open types). Word's default look is the target: an Office
// accent palette for series, gray axis text, hairline gridlines, a centered
// title band and a legend strip outside the plot.

type Rec = Record<string, unknown>;

const isRecord = (v: unknown): v is Rec => typeof v === "object" && v !== null && !Array.isArray(v);
const num = (v: unknown): number | undefined => (typeof v === "number" ? v : undefined);
const str = (v: unknown): string | undefined => (typeof v === "string" ? v : undefined);

/** Office theme accent palette (Word 2013+ default chart colors), order the
 *  series cycle through. */
const ACCENTS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];

const AXIS_TEXT = "#595959";
const GRID_LINE = "#D9D9D9";
const AXIS_LINE = "#BFBFBF";
const LABEL_PX = 12;
const TITLE_PX = 15;

/** Series fill: the explicit solid fill wins; otherwise the accent cycle. */
function seriesFillOf(series: Rec, index: number): string {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const fill = sp?.fill;
  if (isRecord(fill) && fill.type === "solid") {
    const color = isRecord(fill.color) ? (str(fill.color.hex) ?? fill.color) : str(fill.color);
    if (typeof color === "string") return color.replace("#", "").toUpperCase();
  }
  return ACCENTS[index % ACCENTS.length]!;
}

/** The chart model the painter works from, flattened out of the verbatim
 *  ChartSpaceOptions record. */
interface ChartModel {
  type: string;
  categories: string[];
  series: Rec[];
  /** "clustered" | "stacked" | "percentStacked" (bar/line/area groups). */
  grouping: string;
  markers: boolean;
  title: string | undefined;
  legend: boolean;
  legendPosition: string;
  holeSize: number;
}

function readModel(chart: Rec): ChartModel | undefined {
  const type = str(chart.type);
  if (!type) return undefined;
  const series: Rec[] = Array.isArray(chart.series) ? chart.series.filter(isRecord) : [];
  const categories = Array.isArray(chart.categories)
    ? chart.categories.filter((c): c is string => typeof c === "string")
    : [];
  // A legend shows by default once the chart has more than one series to
  // distinguish (Word's c:autoTitleDeleted-style default; c:legend's absence
  // means the app default, not "off").
  const legend = chart.showLegend === true || (chart.showLegend === undefined && series.length > 1);
  return {
    type,
    categories,
    series,
    grouping: str(chart.grouping) ?? "clustered",
    markers: chart.markers === true,
    title:
      typeof chart.title === "string" && chart.title
        ? chart.title
        : isRecord(chart.title) && typeof chart.title.text === "string"
          ? chart.title.text
          : undefined,
    legend,
    legendPosition: str(chart.legendPosition) ?? "bottom",
    holeSize: num(chart.holeSize) ?? 50,
  };
}

/** Series values as numbers (c:val / c:yVal). Scatter/bubble series read the
 *  same slot through `values` here — the projection hands yValues over. */
function valuesOf(series: Rec): number[] {
  const raw = Array.isArray(series.values)
    ? series.values
    : Array.isArray(series.yValues)
      ? series.yValues
      : [];
  return raw.filter((v): v is number => typeof v === "number");
}

/** Nice axis bounds: a 1/2/5×10^k step covering [min, max], zero-based unless
 *  the data is entirely off zero (Word keeps a zero baseline when it can). */
function niceBounds(min: number, max: number): { min: number; max: number; step: number } {
  if (!Number.isFinite(min) || !Number.isFinite(max) || min === max) {
    return { min: min === max ? min - 1 : min, max: max === min ? max + 1 : max, step: 1 };
  }
  const raw = (max - min) / 5;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = ([1, 2, 5, 10] as const).map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  let lo = Math.floor(min / step) * step;
  let hi = Math.ceil(max / step) * step;
  if (min >= 0 && lo > 0) lo = 0;
  if (max <= 0 && hi < 0) hi = 0;
  return { min: lo, max: hi, step };
}

const fmtTick = (v: number): string =>
  Math.abs(v) >= 1000 ? `${Math.round(v / 100) / 10}k` : `${Math.round(v * 100) / 100}`;

/** Text label, horizontally centered on x (or left/right-anchored by align).
 *  With a fixed maxWidth box Leafer anchors the Text at its top-left corner,
 *  so center/right anchors shift the box to keep x on the anchor edge. */
function label(
  tree: IGroup,
  text: string,
  x: number,
  y: number,
  size: number,
  align: "center" | "left" | "right",
  maxWidth?: number,
): void {
  if (!text) return;
  const bx =
    maxWidth == null || align === "left" ? x : align === "center" ? x - maxWidth / 2 : x - maxWidth;
  tree.add(
    new Text({
      x: bx,
      y,
      text,
      fontSize: size,
      fill: AXIS_TEXT,
      textAlign: align,
      verticalAlign: "middle",
      ...(maxWidth != null ? { width: maxWidth, overflow: "ellipsis" } : {}),
    }),
  );
}

/** Gap between neighboring legend entries in a horizontal row. */
const LEGEND_GAP = 24;

/** The painted width of a label, for centering a legend row — the hidden
 *  canvas the break-row marks measure with (node-safe: without a document,
 *  a per-character estimate). */
let legendCtx: CanvasRenderingContext2D | null | undefined;
function measureLabelWidth(text: string, px: number): number {
  legendCtx ??=
    typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  if (!legendCtx) return text.length * px;
  legendCtx.font = `${px}px sans-serif`;
  return legendCtx.measureText(text).width;
}

/** One straight hairline. */
function segment(
  tree: IGroup,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  color: string,
): void {
  tree.add(
    new LeaferPath({
      path: `M ${x0} ${y0} L ${x1} ${y1}`,
      stroke: color,
      strokeWidth: 1,
    }),
  );
}

/** Scatter x/y pairs (c:xVal / c:yVal) as equal-length number arrays. */
function scatterPairs(series: Rec): { x: number[]; y: number[] } {
  const x = Array.isArray(series.xValues)
    ? series.xValues.filter((v): v is number => typeof v === "number")
    : [];
  const y = valuesOf(series);
  const n = Math.min(x.length, y.length);
  return { x: x.slice(0, n), y: y.slice(0, n) };
}

/** One plot rectangle the series render inside. */
interface PlotBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

// ── value-axis charts (column / bar / line / area) ──

/** A chart sub-element's hit registration — the box and any shape arrive in
 *  chart-local coordinates, the hit table is page-local (the registrar in
 *  {@link paintChartMember} folds the origin in). */
type ElementReg = (part: ChartPartHit, x: number, y: number, width: number, height: number) => void;

function paintValueChart(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const horizontal = model.type === "bar";
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const stacked = model.grouping === "stacked" || model.grouping === "percentStacked";
  // Per-category stacks resolve against the stack total (percentStacked
  // normalizes it to 100); plain clusters use the raw data range.
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 1);
  let bounds: { min: number; max: number; step: number };
  if (stacked) {
    const totals: number[] = [];
    for (let c = 0; c < catCount; c++) {
      const sum = all.reduce((acc, vals) => acc + (vals[c] ?? 0), 0);
      totals.push(model.grouping === "percentStacked" ? 100 : sum);
    }
    bounds = niceBounds(Math.min(0, ...totals), Math.max(0, ...totals));
  } else {
    bounds = niceBounds(Math.min(0, ...flat), Math.max(0, ...flat));
  }
  const toPx = (v: number): number =>
    horizontal
      ? plot.x + ((v - bounds.min) / (bounds.max - bounds.min)) * plot.width
      : plot.y + plot.height - ((v - bounds.min) / (bounds.max - bounds.min)) * plot.height;
  // The px → value inverse the editor's value-drag gesture reads (Excel's
  // drag-a-point editing). percentStacked bars paint shares, not the raw
  // values a drag would write, so they stay fixed.
  const span = bounds.max - bounds.min;
  const valueDrag =
    model.grouping === "percentStacked"
      ? undefined
      : horizontal
        ? { a: bounds.min - (plot.x * span) / plot.width, b: span / plot.width, horizontal: true }
        : {
            a: bounds.min + ((plot.y + plot.height) * span) / plot.height,
            b: -span / plot.height,
          };

  // Value axis ticks + gridlines (the category axis shares its baseline).
  const ticks: number[] = [];
  for (let v = bounds.min; v <= bounds.max + bounds.step / 2; v += bounds.step) ticks.push(v);
  for (const t of ticks) {
    const p = toPx(t);
    if (horizontal) {
      segment(tree, p, plot.y, p, plot.y + plot.height, GRID_LINE);
      label(tree, fmtTick(t), p, plot.y + plot.height + 4, LABEL_PX - 1, "center");
    } else {
      segment(tree, plot.x, p, plot.x + plot.width, p, t === bounds.min ? AXIS_LINE : GRID_LINE);
      label(tree, fmtTick(t), plot.x - 6, p, LABEL_PX - 1, "right");
    }
  }
  // Category axis: one band per category (labels under/left of the baseline).
  const band = (horizontal ? plot.height : plot.width) / catCount;
  const catAxisAt = toPx(Math.max(bounds.min, 0));
  for (let c = 0; c < catCount; c++) {
    const text = model.categories[c] ?? String(c + 1);
    if (horizontal) {
      const y = plot.y + (c + 0.5) * band;
      label(tree, text, plot.x - 6, y, LABEL_PX - 1, "right");
    } else {
      const x = plot.x + (c + 0.5) * band;
      label(tree, text, x, catAxisAt + 4, LABEL_PX - 1, "center", band);
    }
  }

  // Series: bars pack beside each other inside the band, lines and areas run
  // through the band midpoints. Stacking stacks instead.
  const slot = band / (stacked ? 1 : model.series.length);
  model.series.forEach((series, si) => {
    const vals = all[si]!;
    const fill = seriesFillOf(series, si);
    const isLine = model.type === "line";
    const isArea = model.type === "area";
    if (isLine || isArea) {
      const pts = vals
        .map((v, c) => ({ x: plot.x + (c + 0.5) * band, y: toPx(v), c }))
        .filter((p) => p.x >= plot.x && p.x <= plot.x + plot.width);
      if (pts.length === 0) return;
      const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
      if (isArea) {
        const base = toPx(Math.max(bounds.min, 0));
        tree.add(
          new LeaferPath({
            path: `${d} L ${pts[pts.length - 1]!.x} ${base} L ${pts[0]!.x} ${base} Z`,
            fill: `#${fill}`,
            opacity: 0.55,
          }),
        );
      }
      tree.add(
        new LeaferPath({ path: d, stroke: `#${fill}`, strokeWidth: 2, strokeJoin: "round" }),
      );
      if (model.markers) {
        for (const p of pts) {
          tree.add(new Ellipse({ x: p.x - 3, y: p.y - 3, width: 6, height: 6, fill: `#${fill}` }));
        }
      }
      if (reg) {
        // The series shape first, the data points after — the click's
        // topmost-last scan makes a point win over the line it sits on.
        const xs = pts.map((p) => p.x);
        const ys = pts.map((p) => p.y);
        reg(
          {
            series: si,
            shape: {
              kind: "poly",
              pts: pts.map((p) => [p.x, p.y] as [number, number]),
              ...(isArea ? { closed: true } : { width: 10 }),
            },
          },
          Math.min(...xs),
          Math.min(...ys),
          Math.max(...xs) - Math.min(...xs),
          Math.max(...ys) - Math.min(...ys),
        );
        for (const p of pts) reg({ series: si, point: p.c, valueDrag }, p.x - 4, p.y - 4, 8, 8);
      }
      return;
    }
    // Column / bar rectangles.
    vals.forEach((v, c) => {
      const base = toPx(Math.max(bounds.min, 0));
      const p = toPx(v);
      if (stacked) {
        // Word stacks in series order; the running sum is the bar's far edge.
        const below = all.slice(0, si).reduce((acc, prev) => acc + (prev[c] ?? 0), 0);
        const from = toPx(below);
        const to = toPx(below + v);
        const y0 = Math.min(from, to);
        const h = Math.abs(p - toPx(below));
        const bar = horizontal
          ? { x: y0, y: plot.y + c * band, width: h, height: band }
          : { x: plot.x + c * band, y: y0, width: band, height: h };
        tree.add(new Rect({ ...bar, fill: `#${fill}` }));
        reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
        return;
      }
      const offset = si * slot;
      const bar = horizontal
        ? {
            x: Math.min(base, p),
            y: plot.y + c * band + offset + slot * 0.1,
            width: Math.abs(p - base),
            height: slot * 0.8,
          }
        : {
            x: plot.x + c * band + offset + slot * 0.1,
            y: Math.min(base, p),
            width: slot * 0.8,
            height: Math.abs(p - base),
          };
      tree.add(new Rect({ ...bar, fill: `#${fill}` }));
      reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
    });
  });
}

// ── pie / doughnut ──

function paintPie(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series[0];
  if (!series) return;
  const vals = valuesOf(series);
  const total = vals.reduce((a, b) => a + Math.max(0, b), 0);
  if (total <= 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const r = Math.min(plot.width, plot.height) / 2;
  const hole = model.type === "doughnut" ? (r * model.holeSize) / 100 : 0;
  let angle = -90; // Word starts the first slice at 12 o'clock.
  vals.forEach((v, i) => {
    const sweep = (Math.max(0, v) / total) * 360;
    if (sweep <= 0) return;
    const a0 = (angle * Math.PI) / 180;
    const a1 = ((angle + sweep) * Math.PI) / 180;
    angle += sweep;
    if (sweep >= 360) {
      tree.add(
        new Ellipse({
          x: cx - r,
          y: cy - r,
          width: r * 2,
          height: r * 2,
          fill: `#${seriesFillOf(series, i)}`,
        }),
      );
      reg?.({ series: 0, point: i }, cx - r, cy - r, r * 2, r * 2);
      return;
    }
    const large = sweep > 180 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0);
    const y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1);
    const y1 = cy + r * Math.sin(a1);
    const d = hole
      ? (() => {
          const hx0 = cx + hole * Math.cos(a0);
          const hy0 = cy + hole * Math.sin(a0);
          const hx1 = cx + hole * Math.cos(a1);
          const hy1 = cy + hole * Math.sin(a1);
          return (
            `M ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} ` +
            `L ${hx1} ${hy1} A ${hole} ${hole} 0 ${large} 0 ${hx0} ${hy0} Z`
          );
        })()
      : `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`;
    tree.add(
      new LeaferPath({
        path: d,
        fill: `#${seriesFillOf(series, i)}`,
        stroke: "#FFFFFF",
        strokeWidth: 1,
      }),
    );
    reg?.(
      {
        series: 0,
        point: i,
        shape: { kind: "wedge", cx, cy, r, a0, a1, ...(hole ? { hole } : {}) },
      },
      cx - r,
      cy - r,
      r * 2,
      r * 2,
    );
  });
}

// ── scatter ──

function paintScatter(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const pairs = model.series.map(scatterPairs);
  const flat = pairs.flat();
  if (flat.length === 0) return;
  const xs = flat.flatMap((p) => p.x);
  const ys = flat.flatMap((p) => p.y);
  const bx = niceBounds(Math.min(...xs), Math.max(...xs));
  const by = niceBounds(Math.min(0, ...ys), Math.max(0, ...ys));
  const px = (v: number) => plot.x + ((v - bx.min) / (bx.max - bx.min)) * plot.width;
  const py = (v: number) => plot.y + plot.height - ((v - by.min) / (by.max - by.min)) * plot.height;
  for (const t of ticksOf(bx)) {
    segment(tree, px(t), plot.y, px(t), plot.y + plot.height, GRID_LINE);
    label(tree, fmtTick(t), px(t), plot.y + plot.height + 4, LABEL_PX - 1, "center");
  }
  for (const t of ticksOf(by)) {
    segment(tree, plot.x, py(t), plot.x + plot.width, py(t), GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, py(t), LABEL_PX - 1, "right");
  }
  model.series.forEach((series, si) => {
    const fill = seriesFillOf(series, si);
    const pair = pairs[si];
    if (!pair) return;
    pair.x.forEach((xv, pi) => {
      tree.add(
        new Ellipse({
          x: px(xv) - 3.5,
          y: py(pair.y[pi]) - 3.5,
          width: 7,
          height: 7,
          fill: `#${fill}`,
        }),
      );
      reg?.({ series: si, point: pi }, px(xv) - 4.5, py(pair.y[pi]) - 4.5, 9, 9);
    });
  });
}

const ticksOf = (b: { min: number; max: number; step: number }): number[] => {
  const out: number[] = [];
  for (let v = b.min; v <= b.max + b.step / 2; v += b.step) out.push(v);
  return out;
};

// ── placeholder (radar / stock / surface / ofPie / bubble — unmodeled) ──

function paintPlaceholder(tree: IGroup, model: ChartModel, plot: PlotBox): void {
  tree.add(
    new Rect({
      x: plot.x,
      y: plot.y,
      width: plot.width,
      height: plot.height,
      fill: "#F3F3F3",
      stroke: "#C4C4C4",
      strokeWidth: 1,
      strokeAlign: "center",
    }),
  );
  label(
    tree,
    model.type.toUpperCase(),
    plot.x + plot.width / 2,
    plot.y + plot.height / 2,
    LABEL_PX,
    "center",
  );
}

// ── legend / title ──

function paintLegend(tree: IGroup, model: ChartModel, box: PlotBox, reg?: ElementReg): void {
  const vertical = model.legendPosition === "right" || model.legendPosition === "left";
  // A pie's legend lists its categories (Word colors each point individually),
  // not the single series every pie has.
  const pie = model.type === "pie" || model.type === "doughnut";
  const entries = pie
    ? model.categories.map((name, i) => ({
        name,
        fill: model.series[0] ? seriesFillOf(model.series[0]!, i) : ACCENTS[i % ACCENTS.length]!,
        point: i,
      }))
    : model.series.map((series, si) => ({
        name: str(series.name) ?? `系列${si + 1}`,
        fill: seriesFillOf(series, si),
        point: -1,
      }));
  if (vertical) {
    const x = model.legendPosition === "left" ? box.x : box.x + box.width - 84;
    // Word centers the legend block vertically against the plot area.
    const top = box.y + Math.max(0, (box.height - entries.length * 20) / 2);
    entries.forEach((entry, i) => {
      const y = top + i * 20;
      tree.add(new Rect({ x, y: y + 5, width: 10, height: 10, fill: `#${entry.fill}` }));
      label(tree, entry.name, x + 14, y + 10, LABEL_PX - 1, "left", 66);
      reg?.(
        { series: 0, legend: true, ...(entry.point >= 0 ? { point: entry.point } : {}) },
        x,
        y,
        84,
        20,
      );
    });
    return;
  }
  // Word lays a horizontal legend as one row of entries at natural width —
  // swatch then label — centered as a whole in its band, never spread
  // edge-to-edge.
  const y = model.legendPosition === "top" ? box.y : box.y + box.height - 18;
  const widths = entries.map((e) => 14 + measureLabelWidth(e.name, LABEL_PX - 1));
  const total =
    widths.reduce((sum, w) => sum + w, 0) + LEGEND_GAP * Math.max(0, entries.length - 1);
  let x = box.x + Math.max(0, (box.width - total) / 2);
  entries.forEach((entry, i) => {
    const w = widths[i]!;
    tree.add(new Rect({ x, y: y + 6, width: 10, height: 10, fill: `#${entry.fill}` }));
    label(tree, entry.name, x + 14, y + 11, LABEL_PX - 1, "left");
    reg?.(
      { series: 0, legend: true, ...(entry.point >= 0 ? { point: entry.point } : {}) },
      x,
      y,
      w,
      20,
    );
    x += w + LEGEND_GAP;
  });
}

/** A chart sub-element's shape re-based from chart-local to page-local px
 *  (the hit table lives in page coordinates, the painter paints in the
 *  chart's own). */
function offsetShape(shape: ChartPartShape, dx: number, dy: number): ChartPartShape {
  if (shape.kind === "wedge") return { ...shape, cx: shape.cx + dx, cy: shape.cy + dy };
  return { ...shape, pts: shape.pts.map(([px, py]) => [px + dx, py + dy] as [number, number]) };
}

// ── entry ──

/** Paint one chart member: title band, legend strip, then the plot. The
 *  whole chart renders inside the member's extent box. With `hits` the
 *  sub-elements register their click boxes (bars, points, wedges, the
 *  series line, legend entries, the title band) — Word's second-stage chart
 *  selection reads them once the chart is framed. */
export function paintChartMember(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "chart" }>,
  hits?: ChartHitContext,
): void {
  const model = isRecord(m.chart) ? readModel(m.chart) : undefined;
  if (!model || model.series.length === 0) {
    // No model — the honest empty frame the picture branch paints.
    tree.add(
      new Rect({
        x: m.x,
        y: m.y,
        width: m.width,
        height: m.height,
        fill: "#F3F3F3",
        stroke: "#C4C4C4",
        strokeWidth: 1,
        strokeAlign: "center",
      }),
    );
    return;
  }
  const chart = new Group({ x: m.x, y: m.y, width: m.width, height: m.height });
  tree.add(chart);
  const reg: ElementReg | undefined = hits
    ? (part, x, y, width, height) => {
        // The value-drag map reads page-local px but its affine coefficients
        // are chart-local: value = a + b·(page − origin), so the origin rides
        // into the intercept as −b·origin — the same fold the box and the
        // exact shapes get here.
        const drag = part.valueDrag;
        const origin = drag && drag.horizontal ? hits.ox + m.x : hits.oy + m.y;
        const chartPart = drag
          ? { ...part, valueDrag: { ...drag, a: drag.a - drag.b * origin } }
          : part;
        hits.ctx.hitBoxes?.push({
          page: hits.ctx.pageIndex,
          x: hits.ox + m.x + x,
          y: hits.oy + m.y + y,
          width,
          height,
          para: hits.para,
          index: hits.index,
          kind: hits.kind,
          chartPart: chartPart.shape
            ? { ...chartPart, shape: offsetShape(chartPart.shape, hits.ox + m.x, hits.oy + m.y) }
            : chartPart,
          ...(hits.ctx.layer === "behind" ? { behind: true } : {}),
        });
      }
    : undefined;
  let top = 0;
  if (model.title) {
    // No width box: the title renders at its natural size anchored center —
    // Word never truncates a chart title, so nothing here may clip it. The
    // label's verticalAlign lift (~6px of ink above the y anchor) is budgeted
    // into the y so the glyphs clear the chart's top clip (inline charts paint
    // inside a clipped holder box — ink above y=0 is cut).
    label(chart, model.title, m.width / 2, top + 10, TITLE_PX, "center");
    reg?.({ title: true }, 0, 0, m.width, TITLE_PX + 8);
    top += TITLE_PX + 8;
  }
  const pie = model.type === "pie" || model.type === "doughnut";
  // A pie's legend (its categories) shows on the single series too — Word
  // defaults every pie to a legend; other charts need a second series.
  const legendSize =
    model.legend && (pie || model.series.length > 1)
      ? model.legendPosition === "right" || model.legendPosition === "left"
        ? { w: 84, h: 0 }
        : { w: 0, h: 20 }
      : { w: 0, h: 0 };
  const legendBox: PlotBox | undefined =
    legendSize.w || legendSize.h
      ? {
          // The band anchors to the edge it names — a top band sits below the
          // title — and the cross-axis spans the chart.
          x: model.legendPosition === "right" ? m.width - legendSize.w : 0,
          y: model.legendPosition === "bottom" ? m.height - legendSize.h : top,
          width: legendSize.w || m.width,
          height: legendSize.h || m.height,
        }
      : undefined;
  const plot: PlotBox = {
    x: (legendBox && model.legendPosition === "left" ? legendSize.w : 0) + 44,
    y: top + (legendBox && model.legendPosition === "top" ? legendSize.h : 0) + 6,
    width:
      m.width -
      44 -
      10 -
      (legendBox && model.legendPosition === "right" ? legendSize.w : 0) -
      (legendBox && model.legendPosition === "left" ? legendSize.w : 0),
    height:
      m.height -
      top -
      14 -
      (legendBox && model.legendPosition === "bottom" ? legendSize.h : 0) -
      (legendBox && model.legendPosition === "top" ? legendSize.h : 0),
  };
  if (legendBox) paintLegend(chart, model, legendBox, reg);
  switch (model.type) {
    case "column":
    case "bar":
    case "line":
    case "area":
      paintValueChart(chart, model, plot, reg);
      break;
    case "pie":
    case "doughnut":
      paintPie(chart, model, plot, reg);
      break;
    case "scatter":
      paintScatter(chart, model, plot, reg);
      break;
    default:
      paintPlaceholder(chart, model, plot);
  }
}
