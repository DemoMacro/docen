// Word's line numbering (w:lnNumType): body text lines count across the
// flow — tables skip (their rows carry no numbers), suppressed paragraphs
// render uncounted — and every countBy-th line paints its number in the
// margin. The counter is cross-page state (continuous runs through page
// breaks; newPage/newSection reset it), so the stage computes every page's
// labels in ONE pass here and the painter only places them.

import type { LineNumberMark } from "@docen/core";
import type { FlowPage, ProjectedLineNumbers } from "@docen/layout";
import { gridPadOf } from "@docen/layout";

/** Every page's line-number labels, keyed by page index. Pages of sections
 *  without numbering collect an empty list. */
export function computeLineNumbers(
  pages: readonly FlowPage[],
  sections: readonly { lineNumbers?: ProjectedLineNumbers }[],
  sectionOfPage: readonly number[],
): Map<number, LineNumberMark[]> {
  const out = new Map<number, LineNumberMark[]>();
  let counter = 0;
  let numberedSection = -1;
  for (const [pageIndex, page] of pages.entries()) {
    const sectionIndex = sectionOfPage[pageIndex] ?? 0;
    const config = sections[sectionIndex]?.lineNumbers;
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
      for (const item of page.items) {
        if (item.block.kind !== "paragraph" || item.block.suppressLineNumbers) continue;
        for (const line of item.block.lines) {
          counter += 1;
          if (counter % countBy === 0) {
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
