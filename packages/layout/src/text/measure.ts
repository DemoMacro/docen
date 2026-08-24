// Run analysis — script itemization and line-box metrics. Width measurement
// itself lives inside pretext (canvas measureText, the engine that paints);
// what this module adds on top is the OOXML script-slot split: a mixed
// Latin/CJK run is itemized by code point into script segments, each carrying
// its slot's family, so one run's Latin and CJK halves measure (and paint)
// with ascii/eastAsia respectively.

import { isCjkCodeUnit, type FontMetrics, type FontSlots } from "../font";
import type { LayoutTextStyle } from "../layout-doc";

/** One same-script stretch of a run. */
export interface ScriptSegment {
  text: string;
  isCjk: boolean;
}

export interface AnalyzedText {
  /** Script segments, in order — concatenating their texts reproduces the run. */
  readonly segments: ScriptSegment[];
  /** Whether any code point itemized to the CJK script (docGrid snapping). */
  readonly hasCjk: boolean;
  /** The run's natural line height: max over its script segments of
   *  face.normalRatio × sizePx (a line box is as tall as its tallest font). */
  readonly naturalPx: number;
}

const CACHE_LIMIT = 4000;

/** Analyzes and caches runs. One instance per layout pass; the cache key
 *  covers everything a metric depends on, so re-flows and re-sizes pay the
 *  cheap path (same determinism contract the paginator was built on). */
export class TextMeasurer {
  private readonly cache = new Map<string, AnalyzedText>();

  constructor(private readonly metrics: FontMetrics) {}

  clearCache(): void {
    this.cache.clear();
  }

  /** The natural line height of a style's default face (ratio × size) — the
   *  paragraph strut metric when no text run exists to measure. Resolves the
   *  latin slot for script-slotted families (the ¶-mark glyph's usual face). */
  naturalOf(style: LayoutTextStyle): number {
    const family = familyOfSlot(style.family, false);
    return (
      this.metrics.normalRatio({ family, bold: style.bold, italic: style.italic }) * style.sizePx
    );
  }

  analyze(text: string, style: LayoutTextStyle): AnalyzedText {
    const key = `${text} ${familyKey(style.family)} ${style.sizePx} ${style.bold ? "b" : ""}${style.italic ? "i" : ""}`;
    const cached = this.cache.get(key);
    if (cached) return cached;

    const analyzed = this.analyzeUncached(text, style);
    if (this.cache.size >= CACHE_LIMIT) this.cache.clear();
    this.cache.set(key, analyzed);
    return analyzed;
  }

  private analyzeUncached(text: string, style: LayoutTextStyle): AnalyzedText {
    const segments: ScriptSegment[] = [];
    let hasCjk = false;
    let naturalPx = 0;
    let segStart = 0;
    let segIsCjk = false;
    const flush = (end: number): void => {
      if (end <= segStart) return;
      const segment = text.slice(segStart, end);
      segments.push({ text: segment, isCjk: segIsCjk });
      const family = familyOfSlot(style.family, segIsCjk);
      const natural =
        this.metrics.normalRatio({ family, bold: style.bold, italic: style.italic }) * style.sizePx;
      if (natural > naturalPx) naturalPx = natural;
    };
    let i = 0;
    while (i < text.length) {
      const lead = text.charCodeAt(i);
      const width = lead >= 0xd800 && lead <= 0xdbff && i + 1 < text.length ? 2 : 1;
      const isCjk = isCjkCodeUnit(text, i);
      if (i === segStart) {
        segIsCjk = isCjk;
      } else if (isCjk !== segIsCjk) {
        flush(i);
        segStart = i;
        segIsCjk = isCjk;
      }
      if (isCjk) hasCjk = true;
      i += width;
    }
    flush(text.length);
    if (naturalPx === 0) naturalPx = style.sizePx * 1.2;
    return { segments, hasCjk, naturalPx };
  }
}

/** The family one script segment of a style renders in (slot by script). */
export function familyOfSlot(family: string | FontSlots, isCjk: boolean): string {
  if (typeof family === "string") return family;
  return (isCjk ? (family.eastAsia ?? family.latin) : (family.latin ?? family.eastAsia)) ?? "";
}

/** A CSS font shorthand for one script segment — the string pretext measures
 *  with (canvas measureText) and the painter draws with (LeaferJS), so the
 *  two can never drift apart. */
export function cssFontOf(style: LayoutTextStyle, family: string): string {
  const parts: string[] = [];
  if (style.italic) parts.push("italic");
  if (style.bold) parts.push("bold");
  parts.push(`${style.sizePx}px`);
  parts.push(family ? `"${family.replace(/"/g, '\\"')}", serif` : "serif");
  return parts.join(" ");
}

/** A stable cache-key form of a family (slots stringify deterministically
 *  enough — same values, same key; different key order only splits cache
 *  entries, never merges distinct fonts). */
function familyKey(family: string | FontSlots): string {
  return typeof family === "string" ? family : `${family.latin ?? ""}|${family.eastAsia ?? ""}`;
}
