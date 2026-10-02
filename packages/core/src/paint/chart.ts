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
const AXIS_LINE = "#000000";
const LABEL_PX = 12;
const TITLE_PX = 18;

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

/** Any explicit series fill suppresses Office's varied single-series bars. */
function hasExplicitFill(series: Rec): boolean {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const fill = sp?.fill;
  return isRecord(fill) && (fill.type === "solid" || fill.type === "gradient");
}

/** A solid stroke color from a series/marker shape, for combo lines. */
function strokeColorOf(series: Rec, index: number): string {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const outline = isRecord(sp?.outline) ? sp!.outline : undefined;
  const color = isRecord(outline?.color) ? str(outline!.color.hex) : str(outline?.color);
  const fallback = ACCENTS[(index + 5) % ACCENTS.length]!;
  return (typeof color === "string" ? color : fallback).replace("#", "").toUpperCase();
}

function strokePxOf(series: Rec): number {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const outline = isRecord(sp?.outline) ? sp!.outline : undefined;
  const width = num(outline?.width);
  return width == null || width < 1 ? 2 : Math.max(1, width / 9525);
}

/** A series or data-point fill as a renderer paint: explicit solid hex,
 *  native linear gradient, or the Office accent cycle. */
function chartFillOf(
  series: Rec,
  index: number,
  point = index,
  width = 1,
  height = 1,
): string | Record<string, unknown> {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const points = Array.isArray(series.dataPoints) ? series.dataPoints.filter(isRecord) : [];
  const pointProperties = points.find((item) => num(item.index) === point);
  const pointSp = isRecord(pointProperties?.shapeProperties)
    ? pointProperties!.shapeProperties
    : undefined;
  const fill = pointSp?.fill ?? sp?.fill;
  if (isRecord(fill) && fill.type === "solid") {
    const color = isRecord(fill.color) ? (str(fill.color.hex) ?? fill.color) : str(fill.color);
    if (typeof color === "string") return color.replace("#", "").toUpperCase();
  }
  if (isRecord(fill) && fill.type === "gradient") {
    const options = isRecord(fill.options) ? fill.options : fill;
    const stops = Array.isArray(options.stops)
      ? options.stops
          .map((stop) => {
            if (!isRecord(stop)) return undefined;
            const offset = num(stop.position);
            const color = isRecord(stop.color)
              ? (str(stop.color.hex) ?? stop.color)
              : str(stop.color);
            return offset == null || typeof color !== "string"
              ? undefined
              : { offset: offset / 100, color: `#${color.replace("#", "").toUpperCase()}` };
          })
          .filter(Boolean)
      : [];
    if (stops.length > 1) {
      const shade = isRecord(options.shade) ? options.shade : undefined;
      const angle = ((num(shade?.angle) ?? 0) * Math.PI) / 180;
      const length = width * Math.abs(Math.cos(angle)) + height * Math.abs(Math.sin(angle));
      return {
        type: "linear",
        from: {
          x: width / 2 - (Math.cos(angle) * length) / 2,
          y: height / 2 - (Math.sin(angle) * length) / 2,
        },
        to: {
          x: width / 2 + (Math.cos(angle) * length) / 2,
          y: height / 2 + (Math.sin(angle) * length) / 2,
        },
        stops,
      };
    }
  }
  if (isVariedBarChartModel(series, index)) return ACCENTS[point % ACCENTS.length]!;
  return ACCENTS[index % ACCENTS.length]!;
}

function isVariedBarChartModel(series: Rec, index: number): boolean {
  return (
    index === 0 &&
    Array.isArray(series.values) &&
    series.values.length > 1 &&
    hasExplicitFill(series) === false
  );
}

/** Office varies a one-series bar's points through the accent palette, so its
 *  legend names categories rather than the lone series. */
function isVariedBarChart(model: ChartModel): boolean {
  return (
    (model.type === "bar" || model.type === "column") &&
    model.series.length === 1 &&
    model.categories.length > 1 &&
    !hasExplicitFill(model.series[0]!)
  );
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
  /** Doughnut/first-slice start offset, degrees clockwise from 12 o'clock. */
  firstSliceAngle: number;
  /** "standard" | "marker" | "filled" (radar polygons). */
  radarStyle: string;
  /** Bubble size scale, percent (100 = default max radius). */
  bubbleScale: number;
  /** What bubbleSize maps to: "area" (sqrt) or "width" (linear). */
  sizeRepresents: string;
  /** Whether the declared value axis draws its major gridlines. */
  gridlines: boolean;
  /** Cluster-to-cluster gap as a percentage of one bar width. */
  gapWidth: number;
  /** Secondary plot groups (bar+line combos), in OOXML plot order. */
  secondaryGroups: Rec[];
  /** The verbatim axis records; dual value axes resolve combo scales. */
  axes: Rec[];
}

function readModel(chart: Rec): ChartModel | undefined {
  const type = str(chart.type);
  if (!type) return undefined;
  const series: Rec[] = Array.isArray(chart.series) ? chart.series.filter(isRecord) : [];
  const categories = Array.isArray(chart.categories)
    ? chart.categories.filter((c): c is string => typeof c === "string")
    : [];
  const axes = Array.isArray(chart.axes) ? chart.axes.filter(isRecord) : [];
  const secondaryGroups = Array.isArray(chart.secondaryGroups)
    ? chart.secondaryGroups.filter(isRecord)
    : [];
  const valueAxis = axes.find((axis) => axis.kind === "value");
  const legendSeries =
    series.length +
    secondaryGroups.reduce(
      (count, group) =>
        count + (Array.isArray(group.series) ? group.series.filter(isRecord).length : 0),
      0,
    );
  // A legend shows by default once the chart has more than one series to
  // distinguish (Word's c:autoTitleDeleted-style default; c:legend's absence
  // means the app default, not "off").
  const legend = chart.showLegend === true || (chart.showLegend === undefined && legendSeries > 1);
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
    legendPosition: str(chart.legendPosition) ?? "right",
    holeSize: num(chart.holeSize) ?? 50,
    firstSliceAngle: num(chart.firstSliceAngle) ?? 0,
    radarStyle: str(chart.radarStyle) ?? "standard",
    bubbleScale: num(chart.bubbleScale) ?? 100,
    sizeRepresents: str(chart.sizeRepresents) ?? "area",
    gridlines:
      valueAxis?.majorGridlines === undefined
        ? false
        : valueAxis.majorGridlines === true || isRecord(valueAxis.majorGridlines),
    gapWidth: num(chart.gapWidth) ?? 150,
    secondaryGroups,
    axes,
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
  // Office prefers the order of magnitude over the next 2×/5× family for
  // just-over-a-power bounds (4981 → 6000 × 1000, not 6000 × 2000).
  const step = raw <= 5 * mag ? mag : 10 * mag;
  let lo = Math.floor(min / step) * step;
  // Excel leaves an integer maximum on the next step (data max 5 -> axis
  // max 6); a fractional value still rounds up to the covering tick.
  const maxSpan = max - Math.floor(max) === 0 ? max + 1 : max;
  // An automatic scale leaves room above the largest positive value before
  // rounding to the major unit; 4981 → 6000 and 5673 → 7000.
  const headroom = max > 0 ? max * 1.1 : maxSpan;
  let hi = Math.ceil(headroom / step) * step;
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
  color: string = AXIS_TEXT,
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
      ...(color === AXIS_TEXT ? {} : { fill: color }),
      textAlign: align,
      verticalAlign: "middle",
      ...(maxWidth != null ? { width: maxWidth, overflow: "ellipsis" } : {}),
    }),
  );
}

/** Gap between neighboring legend entries in a horizontal row. */
const LEGEND_GAP = 24;

/** A right/left legend's actual width: swatch, label gap and measured text,
 *  clamped so long series names cannot consume the plot. */
function legendWidthOf(model: ChartModel): number {
  const names = isVariedBarChart(model)
    ? model.categories
    : [
        ...model.series.map((series, si) => str(series.name) ?? `系列${si + 1}`),
        ...model.secondaryGroups.flatMap((group) =>
          (Array.isArray(group.series) ? group.series.filter(isRecord) : []).map(
            (series, si) => str(series.name) ?? `系列${si + 1}`,
          ),
        ),
      ];
  const textWidth = Math.max(0, ...names.map((name) => measureLabelWidth(name, LABEL_PX - 1)));
  return Math.min(96, Math.max(40, textWidth + 28));
}

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
  width = 1,
): void {
  tree.add(
    new LeaferPath({
      path: `M ${x0} ${y0} L ${x1} ${y1}`,
      stroke: color,
      strokeWidth: width,
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

/** A series' enabled value labels. OOXML defaults every flag off. */
function showsValueLabels(series: Rec): boolean {
  const labels = isRecord(series.dataLabels) ? series.dataLabels : undefined;
  return labels?.showVal === true;
}

/** A label's resolved fill after the PPTX projection theme pass; Word falls
 *  back to axis gray when the label carries no explicit text fill. */
function labelColorOf(series: Rec): string {
  const labels = isRecord(series.dataLabels) ? series.dataLabels : undefined;
  const text = isRecord(labels?.textProperties) ? labels!.textProperties : undefined;
  const paragraphs = Array.isArray(text?.paragraphs) ? text!.paragraphs.filter(isRecord) : [];
  const props = paragraphs
    .map((paragraph) => (isRecord(paragraph.properties) ? paragraph.properties : undefined))
    .find(Boolean);
  const run = isRecord(props?.defaultRunProperties) ? props!.defaultRunProperties : undefined;
  const fill = isRecord(run?.fill) ? run!.fill : undefined;
  const color = isRecord(fill?.color) ? (str(fill!.color.hex) ?? fill!.color) : str(fill?.color);
  return typeof color === "string" ? color.replace("#", "").toUpperCase() : AXIS_TEXT;
}

/** Values in chart-local units; percent labels are decimals in the model. */
function formatChartValue(value: number, format: string | undefined): string {
  if (format?.includes("%")) return `${Math.round(value * 10000) / 100}%`;
  return `${Math.round(value * 100) / 100}`;
}

/** A value-axis record by its OOXML axis id. */
function axisById(model: ChartModel, id: unknown): Rec | undefined {
  return model.axes.find((axis) => num(axis.id) === num(id));
}

/** Explicit axis bounds honor the declared major unit; auto axes use the
 *  painter's 1/2/5 covering scale. */
function valueBoundsOf(values: number[], axis: Rec | undefined) {
  const unit = num(axis?.majorUnit);
  if (!unit || unit <= 0) return niceBounds(Math.min(0, ...values), Math.max(0, ...values));
  const min = Math.min(0, ...values);
  const max = Math.max(0, ...values);
  const lower = Math.floor(min / unit) * unit;
  const upper = Math.ceil(max / unit) * unit;
  // Office leaves one major division of headroom on a signed axis; this also
  // keeps dual combo axes on the same tick count as the primary scale.
  return {
    min: lower,
    max: upper + unit,
    step: unit,
  };
}

/** A value-label's screen point for OOXML's bar label positions. */
function barLabelPoint(
  value: number,
  base: number,
  x: number,
  y: number,
  horizontal: boolean,
): { x: number; y: number } {
  const direction = value >= 0 ? 1 : -1;
  return horizontal ? { x: x + direction * 8, y } : { x, y: y - direction * 8 };
}

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
  const primaryValueAxis = model.axes.find(
    (axis) =>
      axis.kind === "value" &&
      axis.delete !== true &&
      (axis.position === "left" || axis.position === "bottom"),
  );
  let bounds: { min: number; max: number; step: number };
  if (stacked) {
    const totals: number[] = [];
    for (let c = 0; c < catCount; c++) {
      const sum = all.reduce((acc, vals) => acc + (vals[c] ?? 0), 0);
      totals.push(model.grouping === "percentStacked" ? 100 : sum);
    }
    bounds = niceBounds(Math.min(0, ...totals), Math.max(0, ...totals));
  } else {
    bounds = valueBoundsOf(flat, primaryValueAxis);
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
      if (model.gridlines) segment(tree, p, plot.y, p, plot.y + plot.height, GRID_LINE);
      label(tree, fmtTick(t), p, plot.y + plot.height + 4, LABEL_PX - 1, "center");
    } else {
      if (model.gridlines) segment(tree, plot.x, p, plot.x + plot.width, p, GRID_LINE);
      label(
        tree,
        formatChartValue(t, str(primaryValueAxis?.numberFormat)),
        plot.x - 6,
        p,
        LABEL_PX - 1,
        "right",
      );
    }
  }
  // Category axis: one band per category (labels under/left of the baseline).
  const band = (horizontal ? plot.height : plot.width) / catCount;
  const catAxisAt = toPx(Math.max(bounds.min, 0));
  segment(tree, plot.x, catAxisAt, plot.x + plot.width, catAxisAt, AXIS_LINE);
  for (const t of ticks) {
    const p = toPx(t);
    if (horizontal) {
      segment(tree, p, plot.y + plot.height, p, plot.y + plot.height + 4, AXIS_LINE);
    } else {
      segment(tree, plot.x - 4, p, plot.x, p, AXIS_LINE);
    }
  }
  for (let c = 0; c < catCount; c++) {
    const center = (c + 0.5) * band;
    if (horizontal) {
      segment(
        tree,
        plot.x + plot.width,
        plot.y + center,
        plot.x + plot.width + 4,
        plot.y + center,
        AXIS_LINE,
      );
    } else {
      segment(
        tree,
        plot.x + center,
        plot.y + plot.height,
        plot.x + center,
        plot.y + plot.height + 4,
        AXIS_LINE,
      );
    }
    const text = model.categories[c] ?? String(c + 1);
    if (horizontal) {
      const y = plot.y + (c + 0.5) * band;
      label(tree, text, plot.x - 6, y, LABEL_PX - 1, "right");
    } else {
      const x = plot.x + (c + 0.5) * band;
      label(tree, text, x, plot.y + plot.height + 10, LABEL_PX - 1, "center", band);
    }
  }

  // Series: bars pack beside each other inside the band, lines and areas run
  // through the band midpoints. Stacking stacks instead.
  // OOXML's gapWidth reserves space between category clusters; Excel spreads
  // each cluster across the remainder instead of padding every bar separately.
  const slot = band / (stacked ? 1 : model.series.length + model.gapWidth / 100);
  model.series.forEach((series, si) => {
    const vals = all[si]!;
    const fill = seriesFillOf(series, si);
    const pointFill = (point: number, width = 1, height = 1): string | Record<string, unknown> =>
      chartFillOf(series, si, point, width, height);
    const isLine = model.type === "line";
    const isArea = model.type === "area";
    const lineColor = isLine ? strokeColorOf(series, si) : fill;
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
        new LeaferPath({
          path: d,
          ...(isArea ? {} : { fill: "transparent" }),
          stroke: `#${lineColor}`,
          strokeWidth: strokePxOf(series),
          strokeJoin: "round",
        }),
      );
      if (model.markers) {
        for (const p of pts) {
          tree.add(
            new Ellipse({ x: p.x - 3, y: p.y - 3, width: 6, height: 6, fill: `#${lineColor}` }),
          );
        }
      }
      if (showsValueLabels(series)) {
        for (const p of pts) {
          label(
            tree,
            formatChartValue(vals[p.c]!, str(series.formatCode)),
            p.x,
            p.y - 8,
            LABEL_PX - 1,
            "center",
            undefined,
            labelColorOf(series),
          );
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
        const paint = pointFill(c, horizontal ? bar.width : 0, horizontal ? 0 : bar.height);
        tree.add(new Rect({ ...bar, fill: typeof paint === "string" ? `#${paint}` : paint }));
        if (showsValueLabels(series)) {
          const anchor = horizontal
            ? barLabelPoint(
                v,
                base,
                bar.x + (v >= 0 ? bar.width : 0),
                bar.y + bar.height / 2,
                horizontal,
              )
            : barLabelPoint(
                v,
                base,
                bar.x + bar.width / 2,
                bar.y + (v >= 0 ? 0 : bar.height),
                horizontal,
              );
          label(
            tree,
            formatChartValue(v, str(series.formatCode)),
            anchor.x,
            anchor.y,
            LABEL_PX - 1,
            "center",
            undefined,
            labelColorOf(series),
          );
        }
        reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
        return;
      }
      const offset = si * slot;
      const bar = horizontal
        ? {
            x: Math.min(base, p),
            y: plot.y + c * band + offset + slot * 0.1,
            width: Math.abs(p - base),
            height: slot,
          }
        : {
            x: plot.x + c * band + offset + slot * 0.1,
            y: Math.min(base, p),
            width: slot,
            height: Math.abs(p - base),
          };
      const paint = pointFill(c, bar.width, bar.height);
      tree.add(new Rect({ ...bar, fill: typeof paint === "string" ? `#${paint}` : paint }));
      if (showsValueLabels(series)) {
        const anchor = barLabelPoint(
          v,
          base,
          horizontal ? bar.x + (v >= 0 ? bar.width : 0) : bar.x + bar.width / 2,
          horizontal ? bar.y + bar.height / 2 : bar.y + (v >= 0 ? 0 : bar.height),
          horizontal,
        );
        label(
          tree,
          formatChartValue(v, str(series.formatCode)),
          anchor.x,
          anchor.y,
          LABEL_PX - 1,
          "center",
          undefined,
          labelColorOf(series),
        );
      }
      reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
    });
  });
}

// ── pie / doughnut ──

function paintPie(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series[0];
  if (!series) return;
  const piePaintOf = (point: number): string | Record<string, unknown> => {
    const paint = chartFillOf(series, 0, point);
    return typeof paint === "string" ? `#${paint}` : paint;
  };
  const vals = valuesOf(series);
  const total = vals.reduce((a, b) => a + Math.max(0, b), 0);
  if (total <= 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const r = Math.min(plot.width, plot.height) / 2;
  const hole = model.type === "doughnut" ? (r * model.holeSize) / 100 : 0;
  let angle = -90 + model.firstSliceAngle;
  vals.forEach((v, i) => {
    const sweep = (Math.max(0, v) / total) * 360;
    if (sweep <= 0) return;
    const a0 = (angle * Math.PI) / 180;
    const a1 = ((angle + sweep) * Math.PI) / 180;
    angle += sweep;
    if (sweep >= 360) {
      const paint = piePaintOf(i);
      tree.add(
        new Ellipse({
          x: cx - r,
          y: cy - r,
          width: r * 2,
          height: r * 2,
          fill: typeof paint === "string" ? `#${paint}` : paint,
        }),
      );
      reg?.({ series: 0, point: i }, cx - r, cy - r, r * 2, r * 2);
      return;
    }
    const paint = piePaintOf(i);
    const start = (a0 * 180) / Math.PI;
    tree.add(
      new Ellipse({
        x: cx - r,
        y: cy - r,
        width: r * 2,
        height: r * 2,
        ...(hole ? { innerRadius: hole / r } : {}),
        ...(sweep < 360 ? { startAngle: start, endAngle: start + sweep } : {}),
        fill: paint,
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
  xyGrid(tree, bx, by, plot, px, py);
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

/** The scatter/bubble value grid: one hairline + tick label per nice step
 *  on both axes. */
function xyGrid(
  tree: IGroup,
  bx: { min: number; max: number; step: number },
  by: { min: number; max: number; step: number },
  plot: PlotBox,
  px: (v: number) => number,
  py: (v: number) => number,
): void {
  for (const t of ticksOf(bx)) {
    segment(tree, px(t), plot.y, px(t), plot.y + plot.height, GRID_LINE);
    label(tree, fmtTick(t), px(t), plot.y + plot.height + 4, LABEL_PX - 1, "center");
  }
  for (const t of ticksOf(by)) {
    segment(tree, plot.x, py(t), plot.x + plot.width, py(t), GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, py(t), LABEL_PX - 1, "right");
  }
}

// ── radar ──

/** Radar (c:radar): a polygonal web — one vertex per category on the
 *  inscribed circle, series as closed polygons over it (Word's web lines).
 *  Value maps linearly to the radius; a vertex carries a radial value-drag
 *  (Excel drags the vertex along its spoke). */
function paintRadar(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 3);
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const radius = Math.min(plot.width, plot.height) / 2;
  const bounds = niceBounds(Math.min(0, ...flat), Math.max(0, ...flat));
  const span = bounds.max - bounds.min;
  const pointAt = (c: number, v: number): [number, number] => {
    const angle = (Math.PI * 2 * c) / catCount - Math.PI / 2;
    const r = (Math.max(0, v - bounds.min) / span) * radius;
    return [cx + r * Math.cos(angle), cy + r * Math.sin(angle)];
  };
  // Concentric polygon rings per tick; spokes from the center to the outer
  // ring's vertices; category labels just outside those vertices.
  for (const t of ticksOf(bounds)) {
    const ring = Array.from({ length: catCount }, (_, c) => pointAt(c, t));
    tree.add(
      new LeaferPath({
        path: `M ${ring.map(([x, y]) => `${x} ${y}`).join(" L ")} Z`,
        stroke: GRID_LINE,
        strokeWidth: 1,
      }),
    );
  }
  const outer = Array.from({ length: catCount }, (_, c) => {
    const angle = (Math.PI * 2 * c) / catCount - Math.PI / 2;
    const [x, y] = pointAt(c, bounds.max);
    return { x, y, angle };
  });
  for (const v of outer) segment(tree, cx, cy, v.x, v.y, GRID_LINE);
  const drag = { a: bounds.min, b: span / radius, radial: { cx, cy } };
  model.series.forEach((series, si) => {
    const vals = all[si];
    if (!vals) return;
    const pts = Array.from({ length: catCount }, (_, c) => pointAt(c, vals[c] ?? bounds.min));
    const d = `M ${pts.map(([x, y]) => `${x} ${y}`).join(" L ")} Z`;
    const fill = seriesFillOf(series, si);
    if (model.radarStyle === "filled") {
      tree.add(new LeaferPath({ path: d, fill: `#${fill}`, opacity: 0.55 }));
    } else {
      tree.add(
        new LeaferPath({ path: d, stroke: `#${fill}`, strokeWidth: 2, strokeJoin: "round" }),
      );
      if (model.radarStyle === "marker") {
        for (const [x, y] of pts) {
          tree.add(new Ellipse({ x: x - 3, y: y - 3, width: 6, height: 6, fill: `#${fill}` }));
        }
      }
    }
    pts.forEach(([x, y], c) =>
      reg?.({ series: si, point: c, valueDrag: drag }, x - 4, y - 4, 8, 8),
    );
  });
  outer.forEach((v, c) =>
    label(
      tree,
      model.categories[c] ?? String(c + 1),
      v.x + 12 * Math.cos(v.angle),
      v.y + 12 * Math.sin(v.angle),
      LABEL_PX - 1,
      "center",
    ),
  );
}

// ── bubble ──

/** Bubble (c:bubble): scatter points whose radius encodes c:bubbleSize —
 *  proportionally to the value (c:sizeRepresents "width") or to its square
 *  root (the default "area"), the largest at `bubbleScale`% of a quarter of
 *  the plot's short side. No value drag: the radius is a second value, not
 *  the y the pointer would move. */
function paintBubble(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series.map((s, si) => {
    const p = scatterPairs(s);
    // Re-typed from a category chart the series carries values only — index
    // the categories instead of dropping the series.
    if (p.x.length === 0) p.x = p.y.map((_, c) => c);
    const raw = Array.isArray(s.bubbleSize)
      ? s.bubbleSize.filter((v): v is number => typeof v === "number")
      : [];
    const sizes = p.y.map((_, c) => raw[c] ?? 1);
    return { p, sizes, fill: seriesFillOf(s, si) };
  });
  const pts = series.flatMap((s) => s.p.x.map((x, c) => ({ x, y: s.p.y[c], s: s.sizes[c] })));
  if (pts.length === 0) return;
  const xs = pts.map((d) => d.x);
  const ys = pts.map((d) => d.y);
  const bx = niceBounds(Math.min(...xs), Math.max(...xs));
  const by = niceBounds(Math.min(0, ...ys), Math.max(0, ...ys));
  const px = (v: number) => plot.x + ((v - bx.min) / (bx.max - bx.min)) * plot.width;
  const py = (v: number) => plot.y + plot.height - ((v - by.min) / (by.max - by.min)) * plot.height;
  xyGrid(tree, bx, by, plot, px, py);
  const maxSize = Math.max(...pts.map((d) => d.s), 0);
  const maxR = ((Math.min(plot.width, plot.height) / 4) * model.bubbleScale) / 100;
  const radiusOf = (s: number): number =>
    model.sizeRepresents === "width"
      ? (maxR * s) / (maxSize || 1)
      : maxR * Math.sqrt(s / (maxSize || 1));
  series.forEach((s, si) => {
    s.p.x.forEach((xv, c) => {
      const r = radiusOf(s.sizes[c]);
      const x = px(xv);
      const y = py(s.p.y[c]);
      tree.add(
        new Ellipse({
          x: x - r,
          y: y - r,
          width: r * 2,
          height: r * 2,
          fill: `#${s.fill}`,
          opacity: 0.75,
        }),
      );
      reg?.({ series: si, point: c }, x - r - 1, y - r - 1, r * 2 + 2, r * 2 + 2);
    });
  });
}

// ── stock ──

/** Stock (c:stock): the series order is the data slots — open/high/low/close
 *  (four series) or high/low/close (three, no left ticks). Each category
 *  draws a high–low spine with the open/close ticks reaching left/right
 *  (Word's high-low-close and open-high-low-close line stocks). No sub-hits:
 *  the data edits through the Edit Data dialog. */
function paintStock(tree: IGroup, model: ChartModel, plot: PlotBox): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 1);
  const bounds = niceBounds(Math.min(0, ...flat), Math.max(0, ...flat));
  const py = (v: number) =>
    plot.y + plot.height - ((v - bounds.min) / (bounds.max - bounds.min)) * plot.height;
  for (const t of ticksOf(bounds)) {
    segment(tree, plot.x, py(t), plot.x + plot.width, py(t), GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, py(t), LABEL_PX - 1, "right");
  }
  const band = plot.width / catCount;
  const four = all.length >= 4;
  const high = (four ? all[1] : all[0]) ?? [];
  const low = (four ? all[2] : all[1]) ?? [];
  const close = (four ? all[3] : all[2]) ?? [];
  const open = four ? (all[0] ?? []) : undefined;
  for (let c = 0; c < catCount; c++) {
    const cx = plot.x + (c + 0.5) * band;
    const hi = high[c];
    const lo = low[c];
    if (hi == null || lo == null) continue;
    segment(tree, cx, py(hi), cx, py(lo), "#595959", 1.5);
    const tick = band * 0.18;
    if (open && open[c] != null)
      segment(tree, cx - tick, py(open[c]), cx, py(open[c]), "#595959", 1.5);
    if (close[c] != null) segment(tree, cx, py(close[c]), cx + tick, py(close[c]), "#595959", 1.5);
    label(
      tree,
      model.categories[c] ?? String(c + 1),
      cx,
      plot.y + plot.height + 4,
      LABEL_PX - 1,
      "center",
    );
  }
}

/** Secondary combo plots share the primary category bands but use their own
 *  value axis — OOXML's bar+line combo is a scatter group addressed to the
 *  right axis. The line reads the series stroke; its labels use the right
 *  axis' number format. */
function paintSecondaryGroups(tree: IGroup, model: ChartModel, plot: PlotBox): void {
  const catCount = Math.max(model.categories.length, 1);
  const band = plot.width / catCount;
  model.secondaryGroups.forEach((group) => {
    const kind = str(group.type);
    if (kind !== "scatter" && kind !== "line") return;
    const series = Array.isArray(group.series) ? group.series.filter(isRecord) : [];
    const allValues = series.flatMap(valuesOf);
    const axis = axisById(model, Array.isArray(group.axisIds) ? group.axisIds[1] : undefined);
    const bounds = valueBoundsOf(allValues, axis);
    if (allValues.length > 0 && axis?.delete !== true) {
      const py = (value: number) =>
        plot.y + plot.height - ((value - bounds.min) / (bounds.max - bounds.min)) * plot.height;
      for (const tick of ticksOf(bounds)) {
        const y = py(tick);
        segment(tree, plot.x + plot.width, y, plot.x + plot.width + 4, y, AXIS_LINE);
        label(
          tree,
          formatChartValue(tick, str(axis?.numberFormat)),
          plot.x + plot.width + 7,
          y,
          LABEL_PX - 1,
          "left",
        );
      }
    }
    series.forEach((seriesItem) => {
      const values = valuesOf(seriesItem);
      if (values.length === 0) return;
      const py = (value: number) =>
        plot.y + plot.height - ((value - bounds.min) / (bounds.max - bounds.min)) * plot.height;
      const color = strokeColorOf(seriesItem, 0);
      const points = values.map((value, category) => ({
        x: plot.x + (category + 0.5) * band,
        y: py(value),
        category,
        value,
      }));
      const path =
        group.smooth === true && points.length > 2
          ? points
              .map((point, index) => {
                if (index === 0) return `M ${point.x} ${point.y}`;
                const previous = points[index - 1]!;
                const before = points[Math.max(0, index - 2)]!;
                const after = points[Math.min(points.length - 1, index + 1)]!;
                const control1 = {
                  x: previous.x + (point.x - before.x) / 6,
                  y: previous.y + (point.y - before.y) / 6,
                };
                const control2 = {
                  x: point.x - (after.x - previous.x) / 6,
                  y: point.y - (after.y - previous.y) / 6,
                };
                return `C ${control1.x} ${control1.y} ${control2.x} ${control2.y} ${point.x} ${point.y}`;
              })
              .join(" ")
          : points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ");
      tree.add(
        new LeaferPath({
          path,
          fill: "transparent",
          stroke: `#${color}`,
          strokeWidth: strokePxOf(seriesItem),
          strokeCap: "round",
          strokeJoin: "round",
        }),
      );
      const marker = isRecord(seriesItem.marker) ? seriesItem.marker : undefined;
      const showMarkers = group.markers === true || marker !== undefined;
      if (showMarkers) {
        const radius = Math.max(1.5, (num(marker?.size) ?? 5) * 0.7);
        points.forEach((point) => {
          tree.add(
            new Ellipse({
              x: point.x - radius,
              y: point.y - radius,
              width: radius * 2,
              height: radius * 2,
              fill: `#${color}`,
            }),
          );
        });
      }
      if (showsValueLabels(seriesItem)) {
        points.forEach((point) => {
          label(
            tree,
            formatChartValue(point.value, str(axis?.numberFormat)),
            point.x,
            point.y - 8,
            LABEL_PX - 1,
            "center",
            undefined,
            labelColorOf(seriesItem),
          );
        });
      }
    });
  });
}

// ── placeholder (surface / ofPie — unmodeled) ──

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
  // not the single series every pie has. A lone bar series does the same.
  const pie = model.type === "pie" || model.type === "doughnut";
  type LegendEntry = {
    name: string;
    fill: string | Record<string, unknown>;
    line: boolean;
    series: number;
    point?: number;
  };
  const entries: LegendEntry[] =
    pie || isVariedBarChart(model)
      ? model.categories.map((name, i) => ({
          name,
          fill: model.series[0]
            ? chartFillOf(model.series[0]!, 0, i, 10, 10)
            : ACCENTS[i % ACCENTS.length]!,
          line: false,
          series: 0,
          point: i,
        }))
      : [
          ...model.series.map((series, si) => ({
            name: str(series.name) ?? `系列${si + 1}`,
            fill: chartFillOf(series, si, si, 10, 10),
            line: model.type === "line",
            series: si,
          })),
          ...model.secondaryGroups.flatMap((group) =>
            (Array.isArray(group.series) ? group.series.filter(isRecord) : []).map(
              (series, si) => ({
                name: str(series.name) ?? `系列${si + 1}`,
                fill: strokeColorOf(series, si),
                line: true,
                series: model.series.length + si,
              }),
            ),
          ),
        ];
  if (vertical) {
    const width = Math.min(box.width, 96);
    const x = model.legendPosition === "left" ? box.x : box.x + box.width - width;
    // Word centers the legend block vertically against the plot area.
    const top = box.y + Math.max(0, (box.height - entries.length * 20) / 2);
    entries.forEach((entry, i) => {
      const y = top + i * 20;
      if (entry.line) {
        if (typeof entry.fill !== "string") return;
        segment(tree, x + 1, y + 10, x + 9, y + 10, `#${entry.fill}`, 2);
      } else {
        tree.add(
          new Rect({
            x,
            y: y + 5,
            width: 10,
            height: 10,
            fill: typeof entry.fill === "string" ? `#${entry.fill}` : entry.fill,
          }),
        );
      }
      label(tree, entry.name, x + 14, y + 10, LABEL_PX - 1, "left", width - 14);
      reg?.(
        {
          series: entry.series,
          legend: true,
          ...(entry.point != null && entry.point >= 0 ? { point: entry.point } : {}),
        },
        x,
        y,
        width,
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
    if (entry.line) {
      if (typeof entry.fill !== "string") return;
      segment(tree, x + 1, y + 11, x + 9, y + 11, `#${entry.fill}`, 2);
    } else {
      tree.add(
        new Rect({
          x,
          y: y + 6,
          width: 10,
          height: 10,
          fill: typeof entry.fill === "string" ? `#${entry.fill}` : entry.fill,
        }),
      );
    }
    label(tree, entry.name, x + 14, y + 11, LABEL_PX - 1, "left");
    reg?.(
      {
        series: entry.series,
        legend: true,
        ...(entry.point != null && entry.point >= 0 ? { point: entry.point } : {}),
      },
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
        // exact shapes get here. A radial map is translation-invariant in a/b;
        // its center is a position and folds the origin directly.
        const drag = part.valueDrag;
        let chartPart = part;
        if (drag) {
          const dx = hits.ox + m.x;
          const dy = hits.oy + m.y;
          chartPart = drag.radial
            ? {
                ...part,
                valueDrag: {
                  ...drag,
                  radial: { cx: drag.radial.cx + dx, cy: drag.radial.cy + dy },
                },
              }
            : { ...part, valueDrag: { ...drag, a: drag.a - drag.b * (drag.horizontal ? dx : dy) } };
        }
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
    reg?.({ title: true }, 0, 0, m.width, TITLE_PX + 8);
    top += TITLE_PX + 8;
  }
  const pie = model.type === "pie" || model.type === "doughnut";
  // A pie's legend (its categories) shows on the single series too — Word
  // defaults every pie to a legend; varied bar charts do the same.
  const secondarySeriesCount = model.secondaryGroups.reduce(
    (count, group) => count + (Array.isArray(group.series) ? group.series : []).length,
    0,
  );
  const legendSize =
    model.legend &&
    (pie || model.series.length > 1 || secondarySeriesCount > 0 || isVariedBarChart(model))
      ? model.legendPosition === "right" || model.legendPosition === "left"
        ? { w: legendWidthOf(model), h: 0 }
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
  // A pie/doughnut plot is the circle itself; only real title/legend
  // furniture reserves space. Cartesian-chart side margins collapsed a real
  // wide-but-short chart frame into a tiny ring.
  const plot: PlotBox = pie
    ? {
        x: legendBox && model.legendPosition === "left" ? legendSize.w : 0,
        y: top + (legendBox && model.legendPosition === "top" ? legendSize.h : 0),
        width:
          m.width -
          (legendBox && model.legendPosition === "left" ? legendSize.w : 0) -
          (legendBox && model.legendPosition === "right" ? legendSize.w : 0),
        height:
          m.height -
          top -
          (legendBox && model.legendPosition === "bottom" ? legendSize.h : 0) -
          (legendBox && model.legendPosition === "top" ? legendSize.h : 0),
      }
    : {
        x: (legendBox && model.legendPosition === "left" ? legendSize.w : 0) + 76,
        y: top + (legendBox && model.legendPosition === "top" ? legendSize.h : 0) + 16,
        width:
          m.width -
          76 -
          (model.secondaryGroups.length > 0
            ? 46
            : legendBox && model.legendPosition === "right"
              ? 4
              : 10) -
          (legendBox && model.legendPosition === "right" ? legendSize.w : 0) -
          (legendBox && model.legendPosition === "left" ? legendSize.w : 0),
        height:
          m.height -
          top -
          70 -
          (legendBox && model.legendPosition === "bottom" ? legendSize.h : 0) -
          (legendBox && model.legendPosition === "top" ? legendSize.h : 0),
      };
  if (model.title) {
    label(chart, model.title, plot.x + plot.width / 2, top + TITLE_PX / 2 + 2, TITLE_PX, "center");
  }
  if (legendBox) paintLegend(chart, model, legendBox, reg);
  switch (model.type) {
    case "column":
    case "bar":
    case "line":
    case "area":
      paintValueChart(chart, model, plot, reg);
      paintSecondaryGroups(chart, model, plot);
      break;
    case "pie":
    case "doughnut":
      paintPie(chart, model, plot, reg);
      break;
    case "scatter":
      paintScatter(chart, model, plot, reg);
      break;
    case "radar":
      paintRadar(chart, model, plot, reg);
      break;
    case "bubble":
      paintBubble(chart, model, plot, reg);
      break;
    case "stock":
      paintStock(chart, model, plot);
      break;
    default:
      paintPlaceholder(chart, model, plot);
  }
}
