// Run analysis — script itemization and line-box metrics. Width measurement
// itself lives inside pretext (canvas measureText, the engine that paints);
// what this module adds on top is the OOXML script-slot split: a mixed
// Latin/CJK run is itemized by code point into script segments, each carrying
// its slot's family, so one run's Latin and CJK halves measure (and paint)
// with ascii/eastAsia respectively.

import { measureNaturalWidth, prepareWithSegments } from "@docen/pretext";

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

/** The raised/lowered run's size fraction (w:vertAlign — Word's built-in
 *  FootnoteReference style is the same ~65% look). */
const VERT_ALIGN_SCALE = 0.65;
/** The baseline shift as a fraction of the BASE size: superscript raises by
 *  ~1/3 em (CSS `super`), subscript sinks by ~1/6 em. */
export const VERT_ALIGN_RISE = 0.34;
export const VERT_ALIGN_DROP = 0.17;

/** The size a style's glyphs measure and paint at — the base size scaled
 *  down for a raised/lowered run. Measuring (analyze/cssFontOf) and painting
 *  (the renderer's fontSize) must both go through this so the two never
 *  drift. */
export function vertAlignedSizePx(style: LayoutTextStyle): number {
  return style.verticalAlign ? style.sizePx * VERT_ALIGN_SCALE : style.sizePx;
}

/** The baseline offset a raised/lowered run paints at, px (negative = up). */
export function vertAlignBaselineShiftPx(style: LayoutTextStyle): number {
  if (style.verticalAlign === "superscript") return -style.sizePx * VERT_ALIGN_RISE;
  if (style.verticalAlign === "subscript") return style.sizePx * VERT_ALIGN_DROP;
  return 0;
}

/** The alphabetic baseline's offset below a painted Text element's top: the
 *  painter pins each Text's lineHeight to the font size (px form), and
 *  Leafer's baseline formula ((lineHeight + 0.7·fontSize) / 2) then puts the
 *  baseline at 0.85 × fontSize. Consumers that must anchor where the glyphs
 *  actually draw (the caret map's ink band) go through this — if the
 *  painter's lineHeight pin ever changes, this is the one place to update. */
export function leaferBaselinePadPx(fontSize: number): number {
  return 0.85 * fontSize;
}

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
      this.metrics.normalRatio({ family, bold: style.bold, italic: style.italic }) *
      vertAlignedSizePx(style)
    );
  }

  analyze(text: string, style: LayoutTextStyle): AnalyzedText {
    // The key uses the SCALED size: raised/lowered runs of the same base face
    // measure identically, so one cache entry serves both.
    const key = `${text} ${familyKey(style.family)} ${vertAlignedSizePx(style)} ${style.bold ? "b" : ""}${style.italic ? "i" : ""}`;
    const cached = this.cache.get(key);
    if (cached) return cached;

    const analyzed = this.analyzeUncached(text, style);
    // Evict one oldest entry (Map iteration = insertion order) instead of
    // clearing wholesale — a clear wipes the warm working set mid-pass and
    // every subsequent analyze re-pays the full segmentation.
    if (this.cache.size >= CACHE_LIMIT) {
      const oldest = this.cache.keys().next().value;
      if (oldest != null) this.cache.delete(oldest);
    }
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
        this.metrics.normalRatio({ family, bold: style.bold, italic: style.italic }) *
        vertAlignedSizePx(style);
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
    if (naturalPx === 0) naturalPx = vertAlignedSizePx(style) * 1.2;
    return { segments, hasCjk, naturalPx };
  }

  /** One string's advance width — the packer's own canvas measurement (each
   *  script segment in its slot's face, the same fonts a broken line sums),
   *  so a caller-side atom's width never drifts from what the breaker charges
   *  an equivalent run. */
  widthOf(text: string, style: LayoutTextStyle): number {
    const { segments } = this.analyze(text, style);
    let width = 0;
    for (const seg of segments)
      width += measureNaturalWidth(
        prepareWithSegments(seg.text, cssFontOf(style, familyOfSlot(style.family, seg.isCjk)), {
          letterSpacing: style.letterSpacingPx ?? 0,
        }),
      );
    return width;
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
  parts.push(`${vertAlignedSizePx(style)}px`);
  parts.push(family ? `"${family.replace(/"/g, '\\"')}", serif` : "serif");
  return parts.join(" ");
}

/** A stable cache-key form of a family (slots stringify deterministically
 *  enough — same values, same key; different key order only splits cache
 *  entries, never merges distinct fonts). */
function familyKey(family: string | FontSlots): string {
  return typeof family === "string" ? family : `${family.latin ?? ""}|${family.eastAsia ?? ""}`;
}
