// Word's line numbering (w:lnNumType): body text lines count across the
// flow — tables skip (their rows carry no numbers), suppressed paragraphs
// render uncounted — and every countBy-th line paints its number in the
// margin. Two Word display rules ride along: a section-break paragraph
// (the double-rule row) neither counts nor paints, and in a multi-column
// section the later columns' lines still COUNT but paint no mark — Word
// numbers only beside the first column (the left margin), the count keeps
// running so the next numbered line carries the right number. The counter
// is cross-page state (continuous runs through page breaks; newPage/
// newSection reset it), so the stage computes every page's labels in ONE
// pass here and the painter only places them.

import type { LineNumberMark } from "@docen/core";
import type {
  FlowPage,
  ProjectedColumns,
  ProjectedFlowBox,
  ProjectedLineNumbers,
} from "@docen/layout";
import { columnBoxesOf, gridPadOf } from "@docen/layout";

/** The stage's per-section paint inputs this counter reads. */
export interface LineNumberSection {
  flow: ProjectedFlowBox;
  lineNumbers?: ProjectedLineNumbers;
  columns?: ProjectedColumns;
}

/** Every page's line-number labels, keyed by page index. Pages of sections
 *  without numbering collect an empty list. */
export function computeLineNumbers(
  pages: readonly FlowPage[],
  sections: readonly LineNumberSection[],
  sectionOfPage: readonly number[],
): Map<number, LineNumberMark[]> {
  const out = new Map<number, LineNumberMark[]>();
  let counter = 0;
  let numberedSection = -1;
  for (const [pageIndex, page] of pages.entries()) {
    const sectionIndex = sectionOfPage[pageIndex] ?? 0;
    const section = sections[sectionIndex];
    const config = section?.lineNumbers;
    const marks: LineNumberMark[] = [];
    if (config) {
      const start = Math.max(1, config.start);
      const countBy = Math.max(1, config.countBy);
      // newPage resets every page; newSection resets when the numbered
      // section changes; continuous carries across everything.
      if (
        config.restart === "newPage" ||
        (config.restart === "newSection" && sectionIndex !== numberedSection)
      ) {
        counter = 0;
      }
      numberedSection = sectionIndex;
      // The x range of the FIRST column — lines beyond it count unpainted.
      const boxes = columnBoxesOf(section.flow.contentWidthPx, section.columns);
      const firstColRight = boxes.length > 1 ? boxes[0]!.xPx + boxes[0]!.widthPx : null;
      for (const item of page.items) {
        if (item.block.kind !== "paragraph" || item.block.suppressLineNumbers) continue;
        // The section-break paragraph paints the double-rule row — Word
        // shows no number there and the row advances no counter.
        if (item.block.sectionEnd) continue;
        const numbered = firstColRight == null || (item.xPx ?? 0) < firstColRight;
        for (const line of item.block.lines) {
          counter += 1;
          if (numbered && counter % countBy === 0) {
            marks.push({
              yPx: item.yPx + line.yPx + gridPadOf(line),
              num: start + counter - 1,
              sizePx: item.block.markSizePx ?? 12,
            });
          }
        }
      }
    }
    out.set(pageIndex, marks);
  }
  return out;
}
