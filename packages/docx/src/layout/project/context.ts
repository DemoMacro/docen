import type { StylesOptions } from "@office-open/docx";

import type { StyleEntry } from "../../style-cascade";
import type { NumberingIndex } from "./numbering";

/** Per-document projection context, resolved once and threaded down. */
export interface ProjectContext {
  styles: StylesOptions | undefined;
  /** id → character style (w:style type="character") — a run's w:rStyle
   *  resolves its run props here (e.g. "Hyperlink" supplies the blue underline
   *  Word paints body links with, while a TOC entry's un-styled hyperlink
   *  stays plain). */
  characterStyles: Map<string, StyleEntry>;
  numberings: NumberingIndex;
  /** Live list counters per numbering reference (level → count), advanced in
   *  document order as numbered paragraphs project. */
  listCounters: Map<string, number[]>;
  /** Comment ranges open at the current document position (w:commentRangeStart
   *  opened, w:commentRangeEnd not yet seen) — ranges span paragraphs, so the
   *  set lives across the projection walk and every text atom inside tints. */
  openComments: Set<number>;
  /** Footnote id → displayed ordinal, assigned in first-reference order
   *  (Word's numbering: the Nth distinct note referenced shows N; the same id
   *  twice shows the same number). Lives across the whole projection walk. */
  footnoteOrdinals: Map<number, number>;
  /** Endnote id → displayed ordinal — same first-reference-order rule as the
   *  footnotes; painted as lowercase Roman (Word's endnote default numFmt). */
  endnoteOrdinals: Map<number, number>;
}
