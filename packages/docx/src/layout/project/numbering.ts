// Numbering (list) resolution: the reference → levels index built once per
// projection, and the w:numFmt display formatting (decimal, roman, CJK
// numerals, letters) that substitutes a level's %k placeholders.

import { isRecord, measureTwip, num, str, type Rec } from "./guards";

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
const BUILTIN_BULLET_GLYPHS = ["●", "○", "■", "●", "○", "■", "●", "●", "●"];
export const BUILTIN_BULLET_LEVEL = (level: number): NumberingLevel => ({
  format: "bullet",
  text: BUILTIN_BULLET_GLYPHS[Math.min(Math.max(level, 0), 8)],
  leftTw: 720 * (Math.min(Math.max(level, 0), 8) + 1),
  hangingTw: 360,
});

export function indexNumberings(numbering: unknown): NumberingIndex {
  const index: NumberingIndex = new Map();
  if (!isRecord(numbering) || !Array.isArray(numbering.abstractNumberings)) return index;
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
  return index;
}

// ── list-number formats (w:numFmt) ──

const CJK_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const CJK_UNITS = ["", "十", "百", "千"];

/** chineseCounting composition (零 fill between non-zero groups; the 10-19
 *  range drops the leading 一). */
function chineseNumeral(n: number): string {
  if (n < 1 || n > 9999) return String(n);
  const digits: number[] = [];
  for (let rest = n; rest > 0; rest = Math.floor(rest / 10)) digits.unshift(rest % 10);
  let out = "";
  let zeroPending = false;
  digits.forEach((d, i) => {
    const unit = CJK_UNITS[digits.length - 1 - i];
    if (d === 0) {
      if (out) zeroPending = true;
      return;
    }
    if (zeroPending) {
      out += CJK_DIGITS[0];
      zeroPending = false;
    }
    // 10-19 is 十X, not 一十X.
    if (!(d === 1 && unit === "十" && digits.length === 2)) out += CJK_DIGITS[d];
    out += unit;
  });
  return out;
}

const ROMAN_PAIRS: [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

export function romanNumeral(n: number, upper: boolean): string {
  let rest = n;
  let out = "";
  for (const [value, glyph] of ROMAN_PAIRS) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return upper ? out : out.toLowerCase();
}

/** 1→a…26→z, 27→aa (spreadsheet-style, Word's letter numbering). */
function letterNumeral(n: number, upper: boolean): string {
  let out = "";
  let rest = n;
  while (rest > 0) {
    rest--;
    out = String.fromCharCode(97 + (rest % 26)) + out;
    rest = Math.floor(rest / 26);
  }
  return upper ? out.toUpperCase() : out;
}

/** One level's counter under its w:numFmt. Unsupported formats render decimal. */
export function formatListNumber(format: string, n: number): string {
  switch (format) {
    case "lowerLetter":
      return letterNumeral(n, false);
    case "upperLetter":
      return letterNumeral(n, true);
    case "lowerRoman":
      return romanNumeral(n, false);
    case "upperRoman":
      return romanNumeral(n, true);
    case "chineseCounting":
    case "chineseLegalSimplified":
    case "japaneseCounting":
      return chineseNumeral(n);
    default:
      return String(n);
  }
}
