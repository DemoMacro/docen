// Word single-line-height metrics for the faces real documents use most —
// the winAscent/winDescent pair from each face's OS/2 table plus head's
// unitsPerEm, gathered from a Windows font directory. Word derives single
// spacing as winAscent + winDescent + 2 × round(0.15 × (A + D)) over upem
// (SimSun: (220 + 36 + 2×38) / 256 = 1.2969 × size — matches Word exactly);
// CSS `line-height: normal` is a different number per browser, so faces in
// this table skip the browser probe and faces absent from it fall back to
// the probe (an approximation). Bold/italic files of every sampled family
// carry identical metrics, so the key is the family alone — lowercased,
// with the localized aliases OOXML documents actually carry (宋体, MS Mincho
// siblings) pointing at the same triple.

/** One face's vertical metrics, straight from its tables. */
export interface WordFontMetric {
  upem: number;
  winAscent: number;
  winDescent: number;
}

/** Word's single-spacing ratio for one face's tables. */
export function wordLineRatio(m: WordFontMetric): number {
  const sum = m.winAscent + m.winDescent;
  return (sum + 2 * Math.round(0.15 * sum)) / m.upem;
}

// Legacy CJK bitmap-lineage faces (SimSun/SimHei/KaiTi/FangSong and the
// _GB2312 siblings) all share the 256-upem 220/36 triple.
const CJK_LEGACY: WordFontMetric = { upem: 256, winAscent: 220, winDescent: 36 };

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
  "microsoft yahei": { upem: 2048, winAscent: 2080, winDescent: 536 },
  微软雅黑: { upem: 2048, winAscent: 2080, winDescent: 536 },
  "microsoft yahei ui": { upem: 2048, winAscent: 2167, winDescent: 521 },
  dengxian: { upem: 2048, winAscent: 1659, winDescent: 475 },
  等线: { upem: 2048, winAscent: 1659, winDescent: 475 },
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
