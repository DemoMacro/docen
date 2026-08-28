// Font metrics — the engine's vertical-measurement seam. WIDTH measurement
// lives inside pretext (canvas measureText against the browser's own font
// stack — the exact engine that later paints the text, so measure == render
// by construction; fonts are never registered or bundled here). What the
// engine still needs per face is Word's "single line height" ratio — the
// OOXML ADD line-spacing model's multiplier.

import { WORD_FONT_METRICS, wordLineRatio } from "./font-metrics-data";

/** A request for one face's metric. `family` is a resolved CSS-style family
 *  name — OOXML script-slot resolution (eastAsia vs ascii) has already picked
 *  the slot by the time a face is requested. */
export interface FontRequest {
  family: string;
  bold?: boolean;
  italic?: boolean;
}

/** Vertical font metrics. Implementations must be safe to cache and reuse
 *  across layout passes (same input → same output — the determinism the
 *  paginator depends on). */
export interface FontMetrics {
  /** Word's single-line-height ratio for a face: winAscent + winDescent +
   *  2 × round(0.15 × (A + D)) over upem. Faces tabulated in
   *  font-metrics-data resolve to Word's exact number; the browser probe
   *  approximates the rest with CSS `line-height: normal`. */
  normalRatio(request: FontRequest): number;
}

// ── browser implementation (a hidden DOM span probe) ──
// A face tabulated in font-metrics-data returns Word's formula value before
// any probing. Untabulated faces fall back to a span probe: CSS
// `line-height: normal` there is the browser's own metric (close to, but not
// Word's number, and browser-dependent). TextMetrics
// fontBoundingBoxAscent/Descent cover only the glyph boundary box and OMIT
// the font's line-gap, drifting below the rendered `normal`; the span
// measures the full metric the browser actually renders. No DOM (SSR/Node
// tests) → the 1.2 fallback — width measurement still works there via an
// injected canvas (see spec setup).

const FALLBACK_RATIO = 1.2;
/** Fixed probe size: ratio = measuredHeight / PROBE_SIZE_PX stays
 *  font-size-independent (a property of the font, not a size). */
const PROBE_SIZE_PX = 100;

const ratioCache = new Map<string, number>();
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

/** Word's single-line ratio for tabulated faces, else the browser's own
 *  `normal` metric — one value per (family, bold, italic), zero font
 *  registration: the family name resolves through the browser's font stack
 *  (system 宋体, installed Inter, …), the same resolution painting uses. */
export const browserFontMetrics: FontMetrics = {
  normalRatio(request) {
    const key = `${request.family}|${request.bold ? "b" : ""}|${request.italic ? "i" : ""}`;
    const cached = ratioCache.get(key);
    if (cached != null) return cached;
    const word = WORD_FONT_METRICS[request.family.trim().toLowerCase()];
    if (word) {
      const ratio = wordLineRatio(word);
      ratioCache.set(key, ratio);
      return ratio;
    }
    const node = ensureProbe();
    if (!node) return FALLBACK_RATIO;
    const parts: string[] = [];
    if (request.italic) parts.push("italic");
    if (request.bold) parts.push("bold");
    // Empty family → the generic fallback cssFontOf draws with ("serif"), so
    // the metric probes the same face — a bare "100px" font shorthand is
    // invalid CSS and silently no-ops, leaving the probe at the inherited
    // body font (a garbage ratio that collapses every line box).
    parts.push(`${PROBE_SIZE_PX}px`, request.family || "serif");
    let ratio = FALLBACK_RATIO;
    // A detached span reports height 0 — it must be in the layout tree.
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
  },
};

/** Drop the ratio cache. Call after `document.fonts.ready` — fonts loaded
 *  later change the metric, so earlier ratios drift. */
export function clearFontMetricCache(): void {
  ratioCache.clear();
}

// ── OOXML script-slot font resolution ──
// OOXML picks a run's font by the Unicode range of its text (CJK → eastAsia,
// else ascii/hAnsi); `hint` only disambiguates borderline characters. This is
// shared by all three formats (w:rFonts ≡ a:latin/a:ea), so the engine owns
// the slot resolution while adapters map their vocabulary onto the slots.

/** Script-split font slots, the layout projection of `w:rFonts`
 *  (`ascii`/`hAnsi` merge into `latin`) and `a:latin`/`a:ea`. */
export interface FontSlots {
  latin?: string;
  eastAsia?: string;
  /** w:rFonts/@w:hint="eastAsia" — borderline chars resolve to the eastAsia slot. */
  eastAsiaHint?: boolean;
}

/** The CJK block ranges a run's text is tested against (Hangul, CJK
 *  ideographs/extensions, Kana, CJK compatibility, fullwidth forms). */
const CJK_RANGE = /[ᄀ-ᇿ⺀-鿿ꥠ-꥿가-힯豈-﫿぀-ヿ＀-￯]/;

/** Whether a code unit is in the CJK ranges (drives eastAsia slot selection
 *  and docGrid line snapping). */
export function isCjkCodeUnit(unit: string, index: number): boolean {
  return CJK_RANGE.test(unit[index]);
}

/** Whether a single code point (one string element) is in the CJK ranges. */
export function isCjkCodePoint(ch: string): boolean {
  return CJK_RANGE.test(ch);
}

/** Whether any code point of the text falls in the CJK ranges — the
 *  whole-run test behind font-slot resolution. */
export function isCjkText(text: string): boolean {
  return text.length > 0 && CJK_RANGE.test(text);
}

/** Resolve the family a run's text renders in. A string wins as-is; slots
 *  merge over `defaultSlots` (a hint-only run font inherits the defaults,
 *  never replaces them) and the text's script picks the slot. */
export function resolveFontFamily(
  font: string | FontSlots | undefined,
  defaultFont: string | FontSlots | undefined,
  text: string,
): string | null {
  if (typeof font === "string") return font;
  const base = defaultFont && typeof defaultFont === "object" ? defaultFont : {};
  const over = font && typeof font === "object" ? font : {};
  const slots: FontSlots = { ...base, ...over };
  if (isCjkText(text) || slots.eastAsiaHint) {
    return slots.eastAsia ?? slots.latin ?? null;
  }
  return slots.latin ?? slots.eastAsia ?? null;
}
