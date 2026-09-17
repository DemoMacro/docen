// Numbering (list) resolution: the reference → levels index built once per
// projection. The w:numFmt display formatting itself (decimal, roman, CJK
// numerals, letters, kana, …) is layout's numbering-format module — the same
// renderer the page-number field consumes — re-exported here for the
// projection's substitution sites.

import { isRecord, measureTwip, num, str, type Rec } from "./guards";

export { formatNumber as formatListNumber, romanNumeral } from "@docen/layout";

// ── numbering (list) resolution ──

/** One numbering level's layout-relevant fields (w:lvl). */
export interface NumberingLevel {
  format: string;
  text: string;
  leftTw?: number;
  hangingTw?: number;
}

/** reference → levels indexed by w:lvl/@w:ilvl. Bullet levels render today;
 *  numbered formats (decimal…) need a document-order counter — a registered
 *  gap (the projection is a pure per-paragraph walk today). */
export type NumberingIndex = Map<string, NumberingLevel[]>;

/** The built-in bullet list's glyphs and indentation (office-open's
 *  DEFAULT_BULLET_LEVELS, numId 1): a `bullet {level}` paragraph — the sugar
 *  office-open's parser emits for an unresolvable w:numPr, and what a fresh
 *  hand-authored list carries — resolves against this table when no explicit
 *  numbering definition covers it. */
const BUILTIN_BULLET_GLYPHS = Array.from({ length: 9 }, (_, i) => ["●", "○", "■"][i % 3]);
export const BUILTIN_BULLET_LEVEL = (level: number): NumberingLevel => ({
  format: "bullet",
  text: BUILTIN_BULLET_GLYPHS[Math.min(Math.max(level, 0), 8)],
  leftTw: 720 * (Math.min(Math.max(level, 0), 8) + 1),
  hangingTw: 360,
});

// Cache the index by the numbering options reference — the numbering model is
// stable for a document's lifetime, while projectPageFurniture rebuilds the
// index per section (the same WeakMap memo as indexCharacterStyles).
const numberingIndexCache = new WeakMap<object, NumberingIndex>();

export function indexNumberings(numbering: unknown): NumberingIndex {
  if (!isRecord(numbering)) return new Map();
  const cached = numberingIndexCache.get(numbering);
  if (cached) return cached;
  const index: NumberingIndex = new Map();
  if (!Array.isArray(numbering.abstractNumberings)) return index;
  for (const abs of numbering.abstractNumberings) {
    if (!isRecord(abs)) continue;
    const reference = str(abs.reference);
    const levels: NumberingLevel[] = [];
    if (reference && Array.isArray(abs.levels)) {
      for (const lvl of abs.levels) {
        if (!isRecord(lvl)) continue;
        const ind: Rec =
          isRecord(lvl.paragraph) && isRecord(lvl.paragraph.indent) ? lvl.paragraph.indent : {};
        levels[num(lvl.level) ?? 0] = {
          format: typeof lvl.format === "string" ? lvl.format : "bullet",
          text: typeof lvl.text === "string" ? lvl.text : "",
          leftTw: measureTwip(ind.left),
          hangingTw: measureTwip(ind.hanging),
        };
      }
      index.set(reference, levels);
    }
  }
  numberingIndexCache.set(numbering, index);
  return index;
}

// ── list-number formats (w:numFmt) ──
// The formatter lives in @docen/layout (numbering-format.ts) — re-exported
// above alongside the romanNumeral helper the footnote ordinals use.
