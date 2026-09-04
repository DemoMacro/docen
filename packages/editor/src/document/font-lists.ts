// Font data shared by the ribbon font group and the Font dialog — names and
// point sizes are data, not UI copy, so they stay untranslated.

/** Fallback font list shown when the Local Font Access API is unavailable or
 *  denied — includes common CJK faces so a zh host still sees familiar names. */
export const FONT_NAMES = [
  "Microsoft YaHei",
  "Calibri",
  "Arial",
  "Cambria",
  "Times New Roman",
  "Georgia",
  "Verdana",
  "Tahoma",
  "Courier New",
  "Segoe UI",
  "宋体",
  "黑体",
  "楷体",
  "仿宋",
  "等线",
];

/** Chinese size names (Word's zh font-size names) mapped to point values,
 *  largest first (matching the Word zh size picker order). */
export const FONT_SIZES_CN: ReadonlyArray<readonly [string, number]> = [
  ["初号", 42],
  ["小初", 36],
  ["一号", 26],
  ["小一", 24],
  ["二号", 22],
  ["小二", 18],
  ["三号", 16],
  ["小三", 15],
  ["四号", 14],
  ["小四", 12],
  ["五号", 10.5],
  ["小五", 9],
  ["六号", 7.5],
  ["小六", 6.5],
  ["七号", 5.5],
  ["八号", 5],
];

/** Point sizes listed below the Chinese sizes (ascending — Word's zh size
 *  picker orders the numeric sizes small-to-large under the Chinese names). */
export const FONT_SIZES_PT: ReadonlyArray<number> = [
  5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];

/** Word's Underline split menu / Font dialog dropdown — the ST_Underline
 *  patterns (minus "words", which no Word menu exposes) mapped to their i18n
 *  keys; the clear entry ("none") is prepended by each consumer. */
export const UNDERLINE_STYLES: ReadonlyArray<readonly [string, string]> = [
  ["double", "underline-double"],
  ["thick", "underline-thick"],
  ["dotted", "underline-dotted"],
  ["dottedHeavy", "underline-dotted-heavy"],
  ["dash", "underline-dash"],
  ["dashedHeavy", "underline-dashed-heavy"],
  ["dashLong", "underline-dash-long"],
  ["dashLongHeavy", "underline-dash-long-heavy"],
  ["dotDash", "underline-dot-dash"],
  ["dashDotHeavy", "underline-dash-dot-heavy"],
  ["dotDotDash", "underline-dot-dot-dash"],
  ["dashDotDotHeavy", "underline-dash-dot-dot-heavy"],
  ["wave", "underline-wave"],
  ["wavyHeavy", "underline-wavy-heavy"],
  ["wavyDouble", "underline-wavy-double"],
];
