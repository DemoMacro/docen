/**
 * Font `line-height: normal` measurement — the font's true single-line metric
 * (ascent + descent + line-gap), which is exactly Word's "single line height".
 *
 * Why a DOM probe (not canvas measureText): `TextMetrics.fontBoundingBoxAscent
 * /Descent` cover only the glyph boundary box and OMIT the font's line-gap, so
 * they drift below the rendered `normal`. A hidden DOM span with
 * `line-height: normal` measures the full metric the browser actually uses for
 * `normal` (the one the OOXML ADD model — `metric × multiple + linePitch` —
 * needs to match Word).
 *
 * CSS `calc()` cannot reference `normal`, so the editor measures the ratio
 * here (cached per family/bold/italic) and the renderer resolves it from
 * `--docen-font-metric`. The paginator (measure.ts) uses the same ratio so
 * edit == render (C-route invariant).
 */

export interface FontRatioSpec {
  family: string;
  bold: boolean;
  italic: boolean;
}

/** Fallback when the DOM is unavailable (SSR) or measurement fails. */
const FALLBACK_RATIO = 1.2;
/** Fixed probe font-size: ratio = measuredHeight / PROBE_SIZE_PX, so the cached
 *  ratio is font-size-independent (a property of the font, not a size). */
const PROBE_SIZE_PX = 100;

const ratioCache = new Map<string, number>();
/** One reused probe span — created lazily, appended/removed per measurement so
 *  it never participates in layout outside a cache miss. */
let probe: HTMLSpanElement | null = null;

function ensureProbe(): HTMLSpanElement | null {
  if (typeof document === "undefined") return null;
  if (!probe) {
    probe = document.createElement("span");
    probe.style.cssText = "position:absolute;visibility:hidden;white-space:pre;line-height:normal;";
    probe.textContent = "Mg";
  }
  return probe;
}

/** Measure a font's `line-height: normal` ratio (rendered height / font-size),
 *  including the font's line-gap — Word's single-line metric. Cached per
 *  (family, bold, italic). Returns the 1.2 fallback when the DOM is
 *  unavailable (SSR) or the measurement yields nothing. */
export function fontNormalRatio(spec: FontRatioSpec): number {
  const key = `${spec.family}|${spec.bold ? "b" : ""}|${spec.italic ? "i" : ""}`;
  const cached = ratioCache.get(key);
  if (cached != null) return cached;
  const node = ensureProbe();
  if (!node) return FALLBACK_RATIO;
  // Canvas-style font shorthand so the resolved family (incl. the browser's
  // fallback when the named font is absent) matches what the renderer uses.
  const parts: string[] = [];
  if (spec.italic) parts.push("italic");
  if (spec.bold) parts.push("bold");
  parts.push(`${PROBE_SIZE_PX}px`, spec.family);
  let ratio = FALLBACK_RATIO;
  // A detached span reports height 0 — it must be in the layout tree to measure.
  document.body.append(node);
  try {
    node.style.font = parts.join(" ");
    const rect = node.getBoundingClientRect();
    if (rect.height > 0) ratio = rect.height / PROBE_SIZE_PX;
  } finally {
    node.remove();
  }
  ratioCache.set(key, ratio);
  return ratio;
}

/** Drop the ratio cache. Call after `document.fonts.ready` — fonts loaded after
 *  caching change the metric, so a re-measure is needed (paired with
 *  clearMeasureCache on the paginator's prepare cache). */
export function clearFontMetricCache(): void {
  ratioCache.clear();
}
