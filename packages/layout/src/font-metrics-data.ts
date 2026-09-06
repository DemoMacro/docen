// Word single-line-height metrics for the faces real documents use most —
// the winAscent/winDescent pair from each face's OS/2 table plus head's
// unitsPerEm, gathered from a Windows font directory. Word derives single
// spacing from the OS/2 pair alone for Latin cores — COM-measured geometry
// (wdVerticalPositionRelativeToPage pitch, LayoutMode default / single
// spacing): Calibri 11pt 13.35pt = 1.214 ≈ (1950+550)/2048, Arial/Times/
// Tahoma/Segoe UI all match their bare pair within readout precision. CJK
// faces add Word's 15% CJK leading: winAscent + winDescent + 2 × round(0.15
// × (A + D)) over upem — SimSun (220 + 36 + 2×38) / 256 = 1.2969 matches,
// as does Microsoft YaHei (1.66 vs measured ~1.7); see the cjkLeading flag.
// DengXian overrides the formula entirely (wordRatio). CSS `line-height:
// normal` is a different number per browser, so faces in this table skip
// the browser probe and faces absent from it fall back to the probe (an
// approximation). Bold/italic files of every sampled family carry identical
// metrics, so the key is the family alone — lowercased, with the localized
// aliases OOXML documents actually carry (宋体, MS Mincho siblings) pointing
// at the same triple.

/** One face's vertical metrics, straight from its tables. */
export interface WordFontMetric {
  upem: number;
  winAscent: number;
  winDescent: number;
  /** Word's 15% CJK leading applies to this face (single spacing = winAscent
   *  + winDescent + 2 × round(0.15 × (A + D)) over upem). Latin cores omit
   *  it — their Word single spacing is the bare OS/2 pair over upem. */
  cjkLeading?: boolean;
  /** Measured single-spacing ratio overriding the formula — faces whose
   *  rendered Word line box the OS/2 triple does not reproduce (DengXian:
   *  the formula gives 1.3545 but Word renders 1.4×). */
  wordRatio?: number;
}

/** Word's single-spacing ratio for one face's tables. */
export function wordLineRatio(m: WordFontMetric): number {
  if (m.wordRatio != null) return m.wordRatio;
  const sum = m.winAscent + m.winDescent;
  const leading = m.cjkLeading ? 2 * Math.round(0.15 * sum) : 0;
  return (sum + leading) / m.upem;
}

// Legacy CJK bitmap-lineage faces (SimSun/SimHei/KaiTi/FangSong and the
// _GB2312 siblings) all share the 256-upem 220/36 triple.
const CJK_LEGACY: WordFontMetric = {
  upem: 256,
  winAscent: 220,
  winDescent: 36,
  cjkLeading: true,
};

export const WORD_FONT_METRICS: Readonly<Record<string, WordFontMetric>> = {
  simsun: CJK_LEGACY,
  宋体: CJK_LEGACY,
  nsimsun: CJK_LEGACY,
  新宋体: CJK_LEGACY,
  simhei: CJK_LEGACY,
  黑体: CJK_LEGACY,
  kaiti: CJK_LEGACY,
  楷体: CJK_LEGACY,
  kaiti_gb2312: CJK_LEGACY,
  楷体_gb2312: CJK_LEGACY,
  fangsong: CJK_LEGACY,
  仿宋: CJK_LEGACY,
  fangsong_gb2312: CJK_LEGACY,
  仿宋_gb2312: CJK_LEGACY,
  "microsoft yahei": { upem: 2048, winAscent: 2080, winDescent: 536, cjkLeading: true },
  微软雅黑: { upem: 2048, winAscent: 2080, winDescent: 536, cjkLeading: true },
  "microsoft yahei ui": { upem: 2048, winAscent: 2167, winDescent: 521, cjkLeading: true },
  // DengXian: the formula on this true OS/2 triple gives 1.3545, but Word
  // renders 1.4× single spacing (10.5/11/12pt samples agree) — the face's
  // line box includes leading the formula doesn't model, so the measured
  // ratio overrides it.
  dengxian: { upem: 2048, winAscent: 1659, winDescent: 475, wordRatio: 1.4 },
  等线: { upem: 2048, winAscent: 1659, winDescent: 475, wordRatio: 1.4 },
  "times new roman": { upem: 2048, winAscent: 1825, winDescent: 443 },
  arial: { upem: 2048, winAscent: 1854, winDescent: 434 },
  calibri: { upem: 2048, winAscent: 1950, winDescent: 550 },
  "courier new": { upem: 2048, winAscent: 1705, winDescent: 615 },
  cambria: { upem: 2048, winAscent: 1946, winDescent: 455 },
  "segoe ui": { upem: 2048, winAscent: 2210, winDescent: 514 },
  tahoma: { upem: 2048, winAscent: 2049, winDescent: 423 },
  verdana: { upem: 2048, winAscent: 2059, winDescent: 430 },
  georgia: { upem: 2048, winAscent: 1878, winDescent: 449 },
  wingdings: { upem: 2048, winAscent: 1841, winDescent: 432 },
  symbol: { upem: 2048, winAscent: 2059, winDescent: 450 },
};
